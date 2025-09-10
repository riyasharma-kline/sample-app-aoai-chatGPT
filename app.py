import zipfile
import io
import requests
from docx import Document
import copy
import json
import os
import logging
import uuid
import httpx
import asyncio
from quart import (
    Blueprint,
    Quart,
    jsonify,
    make_response,
    request,
    send_from_directory,
    render_template,
    current_app,
)

from openai import AsyncAzureOpenAI
from azure.identity.aio import (
    DefaultAzureCredential,
    get_bearer_token_provider
)
from backend.auth.auth_utils import get_authenticated_user_details
from backend.security.ms_defender_utils import get_msdefender_user_json
from backend.history.cosmosdbservice import CosmosConversationClient
from backend.settings import (
    app_settings,
    MINIMUM_SUPPORTED_AZURE_OPENAI_PREVIEW_API_VERSION
)
from backend.utils import (
    format_as_ndjson,
    format_stream_response,
    format_non_streaming_response,
    convert_to_pf_format,
    format_pf_non_streaming_response,
)

bp = Blueprint("routes", __name__, static_folder="static", template_folder="static")

cosmos_db_ready = asyncio.Event()
import os
import requests
from pptx import Presentation
import tiktoken
from dotenv import load_dotenv
import tempfile
from quart import Blueprint, request, jsonify, send_file
import asyncio
from pptx.util import Pt


load_dotenv()

TOKEN_LIMIT = 3500 
MAX_RETRIES = 3
AZURE_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
DEPLOYMENT_NAME = os.getenv("AZURE_OPENAI_MODEL_NAME")
API_VERSION = os.getenv("AZURE_OPENAI_PREVIEW_API_VERSION")
HEADERS = {
    "Content-Type": "application/json",
    "api-key": os.getenv("AZURE_OPENAI_KEY")
}

GLOSSARY = {
}

def estimate_token_count(text):
    """Estimate the token count of a given text."""
    encoding = tiktoken.get_encoding("cl100k_base") 
    token_count = len(encoding.encode(text))
    return token_count

def apply_glossary(text):
    """Replace predefined terms before translation while handling case variations."""
    for term, translation in GLOSSARY.items():
        text = text.replace(term, translation).replace(term.lower(), translation)
    return text.strip()

def split_into_batches(texts, token_limit):
    """Splits the text into batches while maintaining the token limit."""
    batches = []
    current_batch = []
    current_tokens = 0

    for text in texts:
        tokens = estimate_token_count(text)
        if tokens > token_limit:
            continue
        if current_tokens + tokens > token_limit and current_batch:
            batches.append(current_batch)
            current_batch = []
            current_tokens = 0
        current_batch.append(text)
        current_tokens += tokens

    if current_batch:
        batches.append(current_batch)

    return batches

def translate_text_batch(texts, target_language):
    texts = [text.strip() for text in texts if text.strip()]
    if not texts:
        return texts

    api_url = f"{AZURE_ENDPOINT}/openai/deployments/{DEPLOYMENT_NAME}/chat/completions?api-version={API_VERSION}"
    combined_text = "\n\n".join(texts)

    token_count = estimate_token_count(combined_text)
    if token_count > TOKEN_LIMIT:
        return None

    combined_text = apply_glossary(combined_text)

    body = {
        "messages": [
            {"role": "system", "content": (
                f"You are a professional translator specializing in {target_language}. "
                "Translate the following English text very accurately. "
                "Do not alter numbers, formatting, or structure. Do not translate 'Kline' word. Keep it as it is.\n\n"
            )},
            {"role": "user", "content": f"Translate this into {target_language}:\n{combined_text}"}
        ],
        "max_tokens": 4000,
        "temperature": 0,
        "stream": False,
    }

    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(api_url, headers=HEADERS, json=body, timeout=30)
            data = response.json()
            if response.status_code == 200 and "choices" in data:
                return data["choices"][0]["message"]["content"].split("\n\n")
            else:
                print(f"API Error: {data}")
        except requests.exceptions.RequestException as e:
            print(f"Request Error (Attempt {attempt + 1}): {e}")

    print("Translation failed after retries.")
    return None

def replace_text_in_ref(obj, new_text):
    from pptx.dml.color import RGBColor

    def get_font_color(font):
        try:
            if font.color and hasattr(font.color, "rgb") and font.color.rgb:
                return font.color.rgb
        except Exception:
            pass
        return None

    # Handle table cell or text frame paragraph
    if hasattr(obj, "text_frame"):  # Table cell or shape with text_frame
        tf = obj.text_frame
        # Remove all paragraphs except the first
        while len(tf.paragraphs) > 1:
            tf._element.remove(tf.paragraphs[-1]._element)
        para = tf.paragraphs[0]
        # Save formatting from first run if exists
        if para.runs:
            original_run = para.runs[0]
            font = original_run.font
            size = font.size
            name = font.name
            bold = font.bold
            italic = font.italic
            color = get_font_color(font)
        else:
            size = name = bold = italic = color = None
        bullet_level = para.level
        # Remove all runs and text
        para.clear()
        # Add translated text
        run = para.add_run()
        run.text = new_text
        font = run.font
        if size:
            font.size = size
        if name:
            font.name = name
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic
        if color:
            font.color.rgb = color

        para.level = bullet_level
        if hasattr(para, "bullet"):
            para.bullet = bullet_level is not None
    elif hasattr(obj, "runs"):  # Standalone paragraph
        para = obj
        # Save formatting from first run if exists
        if para.runs:
            original_run = para.runs[0]
            font = original_run.font
            size = font.size
            name = font.name
            bold = font.bold
            italic = font.italic
            color = get_font_color(font)
        else:
            size = name = bold = italic = color = None
        bullet_level = para.level
        # Remove all runs and text
        para.clear()
        # Add translated text
        run = para.add_run()
        run.text = new_text
        font = run.font
        if size:
            font.size = size
        if name:
            font.name = name
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic
        if color:
            font.color.rgb = color

        para.level = bullet_level
        if hasattr(para, "bullet"):
            para.bullet = bullet_level is not None

def collect_text_items_from_shape(shape, text_items):
    if hasattr(shape, "shapes"):
        for subshape in shape.shapes:
            collect_text_items_from_shape(subshape, text_items)
    elif hasattr(shape, "text_frame") and shape.text_frame is not None:
        for paragraph in shape.text_frame.paragraphs:
            text = " ".join(run.text for run in paragraph.runs).strip()
            if text:
                text_items.append((text, paragraph))
    elif hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    text_items.append((cell_text, cell))
def translate_pptx(input_pptx, output_pptx, target_language):
    prs = Presentation(input_pptx)
    skipped_slides = []

    for slide_index, slide in enumerate(prs.slides):
        # print(f"\nProcessing Slide {slide_index + 1}")
        text_items = [] 

        # Collect all translatable text chunks from slide
        for shape in slide.shapes:
            collect_text_items_from_shape(shape, text_items)

        # Translate in small batches and replace immediately
        batch = []
        current_tokens = 0

        for text, ref in text_items:
            tokens = estimate_token_count(text)
            if tokens > TOKEN_LIMIT:
                continue
            if current_tokens + tokens > TOKEN_LIMIT:
                translated = translate_text_batch([t for t, _ in batch], target_language)
                if translated:
                    for (_, obj), trans in zip(batch, translated):
                        replace_text_in_ref(obj, trans)
                else:
                    skipped_slides.append(slide_index + 1)

                batch = []
                current_tokens = 0

            batch.append((text, ref))
            current_tokens += tokens

        if batch:
            translated = translate_text_batch([t for t, _ in batch], target_language)
            if translated:
                for (_, obj), trans in zip(batch, translated):
                    replace_text_in_ref(obj, trans)
            else:
                skipped_slides.append(slide_index + 1)

    prs.save(output_pptx)
    return skipped_slides

@bp.route("/translate", methods=["POST"])
async def translate():
    """Handle translation of uploaded PPTX files."""
    try:
        form = await request.form
        file = (await request.files).get("file")
        target_language = form.get("language")

        if not file or not target_language:
            return jsonify({"error": "Missing file or language"}), 400

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            await file.save(temp_input.name)
            input_path = temp_input.name

        output_path = input_path.replace(".pptx", f"_{target_language}.pptx")

        loop = asyncio.get_event_loop()

        # Translate PPTX (split into batches to avoid exceeding token limits)
        await loop.run_in_executor(
            None, translate_pptx, input_path, output_path, target_language
        )
    
        return await send_file(output_path, as_attachment=True)

    except Exception as e:
        print(f"❌ Translation Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)


@bp.route("/translate-images", methods=["POST"])
async def translate_images():
    """Handle translation of uploaded images, return zip of docx files."""
    try:
        form = await request.form
        files = (await request.files).getlist("images")
        target_language = form.get("language")
        if not files or not target_language:
            return jsonify({"error": "Missing images or language"}), 400

        temp_dir = tempfile.mkdtemp()
        docx_paths = []
        for file in files:
            # Sanitize filename to avoid subdirectory issues
            safe_filename = os.path.basename(file.filename)
            img_path = os.path.join(temp_dir, safe_filename)
            await file.save(img_path)

            with open(img_path, "rb") as img_file:
                img_bytes = img_file.read()
            api_url = f"{AZURE_ENDPOINT}/openai/deployments/{DEPLOYMENT_NAME}/chat/completions?api-version={API_VERSION}"
            import base64
            from imghdr import what
            img_format = what(img_path)
            if img_format == "jpeg":
                mime_type = "image/jpeg"
            elif img_format == "png":
                mime_type = "image/png"
            else:
                mime_type = "image/png" 

            img_b64 = base64.b64encode(img_bytes).decode("utf-8")
            # Removing any accidental line breaks
            img_b64 = img_b64.replace("\n", "")
            body = {
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            f"You are a professional translator specializing in {target_language}. "
                            f"If all the visible text in the image is already in {target_language}, DO NOT change or rewrite it. "
                            f"Just extract and return the text as-is. Otherwise, translate all visible text to {target_language}."
                        )
                    },
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": f"Translate all visible text in this image to {target_language}. Give back only the result text without any additional information."},
                            {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{img_b64}"}}
                        ]
                    }
                ],
                "max_tokens": 4000,
                "temperature": 0,
                "stream": False,
            }
            try:
                response = requests.post(api_url, headers=HEADERS, json=body, timeout=60)
                data = response.json()
                if response.status_code == 200 and "choices" in data:
                    translated_text = data["choices"][0]["message"]["content"]
                else:
                    logging.error(f"OpenAI Vision API error for {safe_filename}: {data}")
                    translated_text = "Error: Could not extract/translate text."
            except Exception as e:
                logging.error(f"OpenAI Vision API request error for {safe_filename}: {e}")
                translated_text = "Error: Could not extract/translate text."

            # Creating docx file
            doc = Document()
            doc.add_paragraph(translated_text)
            docx_name = os.path.splitext(safe_filename)[0] + f"_{target_language}.docx"
            docx_path = os.path.join(temp_dir, docx_name)
            doc.save(docx_path)
            docx_paths.append((docx_name, docx_path))

        # Zipping all docx files
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for docx_name, docx_path in docx_paths:
                zipf.write(docx_path, arcname=docx_name)
        zip_buffer.seek(0)

        # Cleanup temp files to save space
        for _, docx_path in docx_paths:
            if os.path.exists(docx_path):
                os.remove(docx_path)
        for file in files:
            safe_filename = os.path.basename(file.filename)
            img_path = os.path.join(temp_dir, safe_filename)
            if os.path.exists(img_path):
                os.remove(img_path)
        os.rmdir(temp_dir)

        return await send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            attachment_filename=f"translated_images_{target_language}.zip"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@bp.route("/tidy", methods=["POST"])
async def tidy():
    """Handle tidying of uploaded PPTX files."""
    try:
        form = await request.form
        file = (await request.files).get("file")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_input:
            await file.save(temp_input.name)
            input_path = temp_input.name

        output_path = input_path.replace(".pptx", f"_tidied.pptx")

        loop = asyncio.get_event_loop()

        # Translate PPTX (splitting into batches to avoid exceeding token limits)
        await loop.run_in_executor(
            None, tidy_pptx, input_path, output_path
        )
    
        return await send_file(output_path, as_attachment=True)

    except Exception as e:
        print(f"❌ Tidying Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

    finally:
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)

def tidy_pptx(input_pptx, output_pptx):
    prs = Presentation(input_pptx)
    skipped_slides = []

    for slide_index, slide in enumerate(prs.slides):
        text_items = [] 

        # Collect all translatable text chunks from slide
        for shape in slide.shapes:
            collect_text_items_from_shape(shape, text_items)

        # Translate in small batches and replace immediately
        batch = []
        current_tokens = 0

        for text, ref in text_items:
            tokens = estimate_token_count(text)
            if tokens > TOKEN_LIMIT:
                continue
            if current_tokens + tokens > TOKEN_LIMIT:
                translated = tidy_text_batch([t for t, _ in batch])
                if translated:
                    for (_, obj), trans in zip(batch, translated):
                        replace_text_in_ref(obj, trans)
                else:
                    skipped_slides.append(slide_index + 1)

                batch = []
                current_tokens = 0

            batch.append((text, ref))
            current_tokens += tokens

        if batch:
            translated = tidy_text_batch([t for t, _ in batch])
            if translated:
                for (_, obj), trans in zip(batch, translated):
                    replace_text_in_ref(obj, trans)
            else:
                skipped_slides.append(slide_index + 1)

    prs.save(output_pptx)
    return skipped_slides

def tidy_text_batch(texts):
    texts = [text.strip() for text in texts if text.strip()]
    if not texts:
        return texts

    api_url = f"{AZURE_ENDPOINT}/openai/deployments/{DEPLOYMENT_NAME}/chat/completions?api-version={API_VERSION}"
    combined_text = "\n\n".join(texts)

    token_count = estimate_token_count(combined_text)
    if token_count > TOKEN_LIMIT:
        return None


    body = {
        "messages": [
            {"role": "system", "content": (
                f"You are a professional editor specializing in tidying up text. "
                "Tidy the following text very accurately. Solve any grammatical errors, improve clarity, and enhance overall readability. "
                "Do not alter numbers, formatting, or structure.\n\n"
            )},
            {"role": "user", "content": f"Tidy this text:\n{combined_text}"}
        ],
        "max_tokens": 4000,
        "temperature": 0,
        "stream": False,
    }

    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(api_url, headers=HEADERS, json=body, timeout=30)
            data = response.json()
            if response.status_code == 200 and "choices" in data:
                return data["choices"][0]["message"]["content"].split("\n\n")
            else:
                print(f"API Error: {data}")
        except requests.exceptions.RequestException as e:
            print(f"Request Error (Attempt {attempt + 1}): {e}")

    print("Tidying failed after retries.")
    return None

def create_app():
    app = Quart(__name__)
    app.register_blueprint(bp)
    app.config["TEMPLATES_AUTO_RELOAD"] = True
    app.config['MAX_CONTENT_LENGTH'] = 80* 1024 * 1024  # Max input file limit is 80 MB
    
    @app.before_serving
    async def init():
        try:
            app.cosmos_conversation_client = await init_cosmosdb_client()
            cosmos_db_ready.set()
        except Exception as e:
            logging.exception("Failed to initialize CosmosDB client")
            app.cosmos_conversation_client = None
            raise e
    
    return app


@bp.route("/")
async def index():
    return await render_template(
        "index.html",
        title=app_settings.ui.title,
        favicon=app_settings.ui.favicon
    )


@bp.route("/favicon.ico")
async def favicon():
    return await bp.send_static_file("favicon.ico")


@bp.route("/assets/<path:path>")
async def assets(path):
    return await send_from_directory("static/assets", path)


# Debug settings
DEBUG = os.environ.get("DEBUG", "false")
if DEBUG.lower() == "true":
    logging.basicConfig(level=logging.DEBUG)

USER_AGENT = "GitHubSampleWebApp/AsyncAzureOpenAI/1.0.0"


# Frontend Settings via Environment Variables
frontend_settings = {
    "auth_enabled": app_settings.base_settings.auth_enabled,
    "feedback_enabled": (
        app_settings.chat_history and
        app_settings.chat_history.enable_feedback
    ),
    "ui": {
        "title": app_settings.ui.title,
        "logo": app_settings.ui.logo,
        "chat_logo": app_settings.ui.chat_logo or app_settings.ui.logo,
        "chat_title": app_settings.ui.chat_title,
        "chat_description": app_settings.ui.chat_description,
        "show_share_button": app_settings.ui.show_share_button,
        "show_chat_history_button": app_settings.ui.show_chat_history_button,
        # Translate Tab
        "translate_tab_enable": app_settings.ui.translate_tab_enable,
        "translate_tab_title": app_settings.ui.translate_tab_title,
        "translate_tab_description_line1": app_settings.ui.translate_tab_description_line1,
        "translate_tab_description_line2": app_settings.ui.translate_tab_description_line2,
        "translate_tab_languages": app_settings.ui.translate_tab_languages,
        
        # Slide Translation Tab
        "translate_tab_slide_limit": app_settings.ui.translate_tab_slide_limit,
        "translate_tab_slide_upload_container_text": app_settings.ui.translate_tab_slide_upload_container_text,
       
        # Image Translation Tab
        "translate_tab_image_upload_limit": app_settings.ui.translate_tab_image_upload_limit,
        "translate_tab_image_upload_container_text": app_settings.ui.translate_tab_image_upload_container_text,
        
        # Tidy Tab
        "tidy_tab_enable": app_settings.ui.tidy_tab_enable,
        "tidy_tab_title": app_settings.ui.tidy_tab_title,
        "tidy_tab_description_line1": app_settings.ui.tidy_tab_description_line1,
        "tidy_tab_description_line2": app_settings.ui.tidy_tab_description_line2,
        "tidy_tab_slide_limit": app_settings.ui.tidy_tab_slide_limit,
        "tidy_tab_slide_upload_container_text": app_settings.ui.tidy_tab_slide_upload_container_text,
        
    },
    "sanitize_answer": app_settings.base_settings.sanitize_answer,
    "oyd_enabled": app_settings.base_settings.datasource_type,
}


# Enable Microsoft Defender for Cloud Integration
MS_DEFENDER_ENABLED = os.environ.get("MS_DEFENDER_ENABLED", "true").lower() == "true"


azure_openai_tools = []
azure_openai_available_tools = []

# Initialize Azure OpenAI Client
async def init_openai_client():
    azure_openai_client = None
    
    try:
        # API version check
        if (
            app_settings.azure_openai.preview_api_version
            < MINIMUM_SUPPORTED_AZURE_OPENAI_PREVIEW_API_VERSION
        ):
            raise ValueError(
                f"The minimum supported Azure OpenAI preview API version is '{MINIMUM_SUPPORTED_AZURE_OPENAI_PREVIEW_API_VERSION}'"
            )

        # Endpoint
        if (
            not app_settings.azure_openai.endpoint and
            not app_settings.azure_openai.resource
        ):
            raise ValueError(
                "AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_RESOURCE is required"
            )

        endpoint = (
            app_settings.azure_openai.endpoint
            if app_settings.azure_openai.endpoint
            else f"https://{app_settings.azure_openai.resource}.openai.azure.com/"
        )

        # Authentication
        aoai_api_key = app_settings.azure_openai.key
        ad_token_provider = None
        if not aoai_api_key:
            logging.debug("No AZURE_OPENAI_KEY found, using Azure Entra ID auth")
            async with DefaultAzureCredential() as credential:
                ad_token_provider = get_bearer_token_provider(
                    credential,
                    "https://cognitiveservices.azure.com/.default"
                )

        # Deployment
        deployment = app_settings.azure_openai.model
        if not deployment:
            raise ValueError("AZURE_OPENAI_MODEL is required")

        # Default Headers
        default_headers = {"x-ms-useragent": USER_AGENT}

        # Remote function calls
        if app_settings.azure_openai.function_call_azure_functions_enabled:
            azure_functions_tools_url = f"{app_settings.azure_openai.function_call_azure_functions_tools_base_url}?code={app_settings.azure_openai.function_call_azure_functions_tools_key}"
            async with httpx.AsyncClient() as client:
                response = await client.get(azure_functions_tools_url)
            response_status_code = response.status_code
            if response_status_code == httpx.codes.OK:
                azure_openai_tools.extend(json.loads(response.text))
                for tool in azure_openai_tools:
                    azure_openai_available_tools.append(tool["function"]["name"])
            else:
                logging.error(f"An error occurred while getting OpenAI Function Call tools metadata: {response.status_code}")

        
        azure_openai_client = AsyncAzureOpenAI(
            api_version=app_settings.azure_openai.preview_api_version,
            api_key=aoai_api_key,
            azure_ad_token_provider=ad_token_provider,
            default_headers=default_headers,
            azure_endpoint=endpoint,
        )

        return azure_openai_client
    except Exception as e:
        logging.exception("Exception in Azure OpenAI initialization", e)
        azure_openai_client = None
        raise e

async def openai_remote_azure_function_call(function_name, function_args):
    if app_settings.azure_openai.function_call_azure_functions_enabled is not True:
        return

    azure_functions_tool_url = f"{app_settings.azure_openai.function_call_azure_functions_tool_base_url}?code={app_settings.azure_openai.function_call_azure_functions_tool_key}"
    headers = {'content-type': 'application/json'}
    body = {
        "tool_name": function_name,
        "tool_arguments": json.loads(function_args)
    }
    async with httpx.AsyncClient() as client:
        response = await client.post(azure_functions_tool_url, data=json.dumps(body), headers=headers)
    response.raise_for_status()

    return response.text

async def init_cosmosdb_client():
    cosmos_conversation_client = None
    if app_settings.chat_history:
        try:
            cosmos_endpoint = (
                f"https://{app_settings.chat_history.account}.documents.azure.com:443/"
            )

            if not app_settings.chat_history.account_key:
                async with DefaultAzureCredential() as cred:
                    credential = cred
                    
            else:
                credential = app_settings.chat_history.account_key

            cosmos_conversation_client = CosmosConversationClient(
                cosmosdb_endpoint=cosmos_endpoint,
                credential=credential,
                database_name=app_settings.chat_history.database,
                container_name=app_settings.chat_history.conversations_container,
                enable_message_feedback=app_settings.chat_history.enable_feedback,
            )
        except Exception as e:
            logging.exception("Exception in CosmosDB initialization", e)
            cosmos_conversation_client = None
            raise e
    else:
        logging.debug("CosmosDB not configured")

    return cosmos_conversation_client


def prepare_model_args(request_body, request_headers):
    request_messages = request_body.get("messages", [])
    messages = []
    if not app_settings.datasource:
        messages = [
            {
                "role": "system",
                "content": app_settings.azure_openai.system_message
            }
        ]

    for message in request_messages:
        if message:
            match message["role"]:
                case "user":
                    messages.append(
                        {
                            "role": message["role"],
                            "content": message["content"]
                        }
                    )
                case "assistant" | "function" | "tool":
                    messages_helper = {}
                    messages_helper["role"] = message["role"]
                    if "name" in message:
                        messages_helper["name"] = message["name"]
                    if "function_call" in message:
                        messages_helper["function_call"] = message["function_call"]
                    messages_helper["content"] = message["content"]
                    if "context" in message:
                        context_obj = json.loads(message["context"])
                        messages_helper["context"] = context_obj
                    
                    messages.append(messages_helper)


    user_security_context = None
    if (MS_DEFENDER_ENABLED):
        authenticated_user_details = get_authenticated_user_details(request_headers)
        application_name = app_settings.ui.title
        user_security_context = get_msdefender_user_json(authenticated_user_details, request_headers, application_name )  # security component introduced here https://learn.microsoft.com/en-us/azure/defender-for-cloud/gain-end-user-context-ai
    

    model_args = {
        "messages": messages,
        "temperature": app_settings.azure_openai.temperature,
        "max_tokens": app_settings.azure_openai.max_tokens,
        "top_p": app_settings.azure_openai.top_p,
        "stop": app_settings.azure_openai.stop_sequence,
        "stream": app_settings.azure_openai.stream,
        "model": app_settings.azure_openai.model
    }

    if len(messages) > 0:
        if messages[-1]["role"] == "user":
            if app_settings.azure_openai.function_call_azure_functions_enabled and len(azure_openai_tools) > 0:
                model_args["tools"] = azure_openai_tools

            if app_settings.datasource:
                model_args["extra_body"] = {
                    "data_sources": [
                        app_settings.datasource.construct_payload_configuration(
                            request=request
                        )
                    ]
                }

    model_args_clean = copy.deepcopy(model_args)
    if model_args_clean.get("extra_body"):
        secret_params = [
            "key",
            "connection_string",
            "embedding_key",
            "encoded_api_key",
            "api_key",
        ]
        for secret_param in secret_params:
            if model_args_clean["extra_body"]["data_sources"][0]["parameters"].get(
                secret_param
            ):
                model_args_clean["extra_body"]["data_sources"][0]["parameters"][
                    secret_param
                ] = "*****"
        authentication = model_args_clean["extra_body"]["data_sources"][0][
            "parameters"
        ].get("authentication", {})
        for field in authentication:
            if field in secret_params:
                model_args_clean["extra_body"]["data_sources"][0]["parameters"][
                    "authentication"
                ][field] = "*****"
        embeddingDependency = model_args_clean["extra_body"]["data_sources"][0][
            "parameters"
        ].get("embedding_dependency", {})
        if "authentication" in embeddingDependency:
            for field in embeddingDependency["authentication"]:
                if field in secret_params:
                    model_args_clean["extra_body"]["data_sources"][0]["parameters"][
                        "embedding_dependency"
                    ]["authentication"][field] = "*****"

    if model_args.get("extra_body") is None:
        model_args["extra_body"] = {}
    if user_security_context:  # security component introduced here https://learn.microsoft.com/en-us/azure/defender-for-cloud/gain-end-user-context-ai     
                model_args["extra_body"]["user_security_context"]= user_security_context.to_dict()
    logging.debug(f"REQUEST BODY: {json.dumps(model_args_clean, indent=4)}")

    return model_args


async def promptflow_request(request):
    try:
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {app_settings.promptflow.api_key}",
        }
        # Adding timeout for scenarios where response takes longer to come back
        logging.debug(f"Setting timeout to {app_settings.promptflow.response_timeout}")
        async with httpx.AsyncClient(
            timeout=float(app_settings.promptflow.response_timeout)
        ) as client:
            pf_formatted_obj = convert_to_pf_format(
                request,
                app_settings.promptflow.request_field_name,
                app_settings.promptflow.response_field_name
            )
            # NOTE: This only support question and chat_history parameters
            # If you need to add more parameters, you need to modify the request body
            response = await client.post(
                app_settings.promptflow.endpoint,
                json={
                    app_settings.promptflow.request_field_name: pf_formatted_obj[-1]["inputs"][app_settings.promptflow.request_field_name],
                    "chat_history": pf_formatted_obj[:-1],
                },
                headers=headers,
            )
        resp = response.json()
        resp["id"] = request["messages"][-1]["id"]
        return resp
    except Exception as e:
        logging.error(f"An error occurred while making promptflow_request: {e}")


async def process_function_call(response):
    response_message = response.choices[0].message
    messages = []

    if response_message.tool_calls:
        for tool_call in response_message.tool_calls:
            # Check if function exists
            if tool_call.function.name not in azure_openai_available_tools:
                continue
            
            function_response = await openai_remote_azure_function_call(tool_call.function.name, tool_call.function.arguments)

            # adding assistant response to messages
            messages.append(
                {
                    "role": response_message.role,
                    "function_call": {
                        "name": tool_call.function.name,
                        "arguments": tool_call.function.arguments,
                    },
                    "content": None,
                }
            )
            
            # adding function response to messages
            messages.append(
                {
                    "role": "function",
                    "name": tool_call.function.name,
                    "content": function_response,
                }
            )  # extend conversation with function response
        
        return messages
    
    return None

async def send_chat_request(request_body, request_headers):
    filtered_messages = []
    messages = request_body.get("messages", [])
    for message in messages:
        if message.get("role") != 'tool':
            filtered_messages.append(message)
            
    request_body['messages'] = filtered_messages
    model_args = prepare_model_args(request_body, request_headers)

    try:
        azure_openai_client = await init_openai_client()
        raw_response = await azure_openai_client.chat.completions.with_raw_response.create(**model_args)
        response = raw_response.parse()
        apim_request_id = raw_response.headers.get("apim-request-id") 
    except Exception as e:
        logging.exception("Exception in send_chat_request")
        raise e

    return response, apim_request_id


async def complete_chat_request(request_body, request_headers):
    if app_settings.base_settings.use_promptflow:
        response = await promptflow_request(request_body)
        history_metadata = request_body.get("history_metadata", {})
        return format_pf_non_streaming_response(
            response,
            history_metadata,
            app_settings.promptflow.response_field_name,
            app_settings.promptflow.citations_field_name
        )
    else:
        response, apim_request_id = await send_chat_request(request_body, request_headers)
        history_metadata = request_body.get("history_metadata", {})
        non_streaming_response = format_non_streaming_response(response, history_metadata, apim_request_id)

        if app_settings.azure_openai.function_call_azure_functions_enabled:
            function_response = await process_function_call(response)  # Add await here

            if function_response:
                request_body["messages"].extend(function_response)

                response, apim_request_id = await send_chat_request(request_body, request_headers)
                history_metadata = request_body.get("history_metadata", {})
                non_streaming_response = format_non_streaming_response(response, history_metadata, apim_request_id)

    return non_streaming_response

class AzureOpenaiFunctionCallStreamState():
    def __init__(self):
        self.tool_calls = []                # All tool calls detected in the stream
        self.tool_name = ""                 # Tool name being streamed
        self.tool_arguments_stream = ""     # Tool arguments being streamed
        self.current_tool_call = None       # JSON with the tool name and arguments currently being streamed
        self.function_messages = []         # All function messages to be appended to the chat history
        self.streaming_state = "INITIAL"    # Streaming state (INITIAL, STREAMING, COMPLETED)


async def process_function_call_stream(completionChunk, function_call_stream_state, request_body, request_headers, history_metadata, apim_request_id):
    if hasattr(completionChunk, "choices") and len(completionChunk.choices) > 0:
        response_message = completionChunk.choices[0].delta
        
        # Function calling stream processing
        if response_message.tool_calls and function_call_stream_state.streaming_state in ["INITIAL", "STREAMING"]:
            function_call_stream_state.streaming_state = "STREAMING"
            for tool_call_chunk in response_message.tool_calls:
                # New tool call
                if tool_call_chunk.id:
                    if function_call_stream_state.current_tool_call:
                        function_call_stream_state.tool_arguments_stream += tool_call_chunk.function.arguments if tool_call_chunk.function.arguments else ""
                        function_call_stream_state.current_tool_call["tool_arguments"] = function_call_stream_state.tool_arguments_stream
                        function_call_stream_state.tool_arguments_stream = ""
                        function_call_stream_state.tool_name = ""
                        function_call_stream_state.tool_calls.append(function_call_stream_state.current_tool_call)

                    function_call_stream_state.current_tool_call = {
                        "tool_id": tool_call_chunk.id,
                        "tool_name": tool_call_chunk.function.name if function_call_stream_state.tool_name == "" else function_call_stream_state.tool_name
                    }
                else:
                    function_call_stream_state.tool_arguments_stream += tool_call_chunk.function.arguments if tool_call_chunk.function.arguments else ""
                
        # Function call - Streaming completed
        elif response_message.tool_calls is None and function_call_stream_state.streaming_state == "STREAMING":
            function_call_stream_state.current_tool_call["tool_arguments"] = function_call_stream_state.tool_arguments_stream
            function_call_stream_state.tool_calls.append(function_call_stream_state.current_tool_call)
            
            for tool_call in function_call_stream_state.tool_calls:
                tool_response = await openai_remote_azure_function_call(tool_call["tool_name"], tool_call["tool_arguments"])

                function_call_stream_state.function_messages.append({
                    "role": "assistant",
                    "function_call": {
                        "name" : tool_call["tool_name"],
                        "arguments": tool_call["tool_arguments"]
                    },
                    "content": None
                })
                function_call_stream_state.function_messages.append({
                    "tool_call_id": tool_call["tool_id"],
                    "role": "function",
                    "name": tool_call["tool_name"],
                    "content": tool_response,
                })
            
            function_call_stream_state.streaming_state = "COMPLETED"
            return function_call_stream_state.streaming_state
        
        else:
            return function_call_stream_state.streaming_state


async def stream_chat_request(request_body, request_headers):
    response, apim_request_id = await send_chat_request(request_body, request_headers)
    history_metadata = request_body.get("history_metadata", {})
    
    async def generate(apim_request_id, history_metadata):
        if app_settings.azure_openai.function_call_azure_functions_enabled:
            # Maintain state during function call streaming
            function_call_stream_state = AzureOpenaiFunctionCallStreamState()
            
            async for completionChunk in response:
                stream_state = await process_function_call_stream(completionChunk, function_call_stream_state, request_body, request_headers, history_metadata, apim_request_id)
                
                # No function call, asistant response
                if stream_state == "INITIAL":
                    yield format_stream_response(completionChunk, history_metadata, apim_request_id)

                # Function call stream completed, functions were executed.
                # Append function calls and results to history and send to OpenAI, to stream the final answer.
                if stream_state == "COMPLETED":
                    request_body["messages"].extend(function_call_stream_state.function_messages)
                    function_response, apim_request_id = await send_chat_request(request_body, request_headers)
                    async for functionCompletionChunk in function_response:
                        yield format_stream_response(functionCompletionChunk, history_metadata, apim_request_id)
                
        else:
            async for completionChunk in response:
                yield format_stream_response(completionChunk, history_metadata, apim_request_id)

    return generate(apim_request_id=apim_request_id, history_metadata=history_metadata)


async def conversation_internal(request_body, request_headers):
    try:
        if app_settings.azure_openai.stream and not app_settings.base_settings.use_promptflow:
            result = await stream_chat_request(request_body, request_headers)
            response = await make_response(format_as_ndjson(result))
            response.timeout = None
            response.mimetype = "application/json-lines"
            return response
        else:
            result = await complete_chat_request(request_body, request_headers)
            return jsonify(result)

    except Exception as ex:
        logging.exception(ex)
        if hasattr(ex, "status_code"):
            return jsonify({"error": str(ex)}), ex.status_code
        else:
            return jsonify({"error": str(ex)}), 500


@bp.route("/conversation", methods=["POST"])
async def conversation():
    if not request.is_json:
        return jsonify({"error": "request must be json"}), 415
    request_json = await request.get_json()

    return await conversation_internal(request_json, request.headers)


@bp.route("/frontend_settings", methods=["GET"])
def get_frontend_settings():
    try:
        return jsonify(frontend_settings), 200
    except Exception as e:
        logging.exception("Exception in /frontend_settings")
        return jsonify({"error": str(e)}), 500


## Conversation History API ##
@bp.route("/history/generate", methods=["POST"])
async def add_conversation():
    await cosmos_db_ready.wait()
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    try:
        # make sure cosmos is configured
        if not current_app.cosmos_conversation_client:
            raise Exception("CosmosDB is not configured or not working")

        # check for the conversation_id, if the conversation is not set, we will create a new one
        history_metadata = {}
        if not conversation_id:
            title = await generate_title(request_json["messages"])
            conversation_dict = await current_app.cosmos_conversation_client.create_conversation(
                user_id=user_id, title=title
            )
            conversation_id = conversation_dict["id"]
            history_metadata["title"] = title
            history_metadata["date"] = conversation_dict["createdAt"]

        ## Format the incoming message object in the "chat/completions" messages format
        ## then write it to the conversation history in cosmos
        messages = request_json["messages"]
        if len(messages) > 0 and messages[-1]["role"] == "user":
            createdMessageValue = await current_app.cosmos_conversation_client.create_message(
                uuid=str(uuid.uuid4()),
                conversation_id=conversation_id,
                user_id=user_id,
                input_message=messages[-1],
            )
            if createdMessageValue == "Conversation not found":
                raise Exception(
                    "Conversation not found for the given conversation ID: "
                    + conversation_id
                    + "."
                )
        else:
            raise Exception("No user message found")

        # Submit request to Chat Completions for response
        request_body = await request.get_json()
        history_metadata["conversation_id"] = conversation_id
        request_body["history_metadata"] = history_metadata
        return await conversation_internal(request_body, request.headers)

    except Exception as e:
        logging.exception("Exception in /history/generate")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/update", methods=["POST"])
async def update_conversation():
    await cosmos_db_ready.wait()
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    try:
        # make sure cosmos is configured
        if not current_app.cosmos_conversation_client:
            raise Exception("CosmosDB is not configured or not working")

        # check for the conversation_id, if the conversation is not set, we will create a new one
        if not conversation_id:
            raise Exception("No conversation_id found")

        ## Format the incoming message object in the "chat/completions" messages format
        ## then write it to the conversation history in cosmos
        messages = request_json["messages"]
        if len(messages) > 0 and messages[-1]["role"] == "assistant":
            if len(messages) > 1 and messages[-2].get("role", None) == "tool":
                # write the tool message first
                await current_app.cosmos_conversation_client.create_message(
                    uuid=str(uuid.uuid4()),
                    conversation_id=conversation_id,
                    user_id=user_id,
                    input_message=messages[-2],
                )
            # write the assistant message
            await current_app.cosmos_conversation_client.create_message(
                uuid=messages[-1]["id"],
                conversation_id=conversation_id,
                user_id=user_id,
                input_message=messages[-1],
            )
        else:
            raise Exception("No bot messages found")

        # Submit request to Chat Completions for response
        response = {"success": True}
        return jsonify(response), 200

    except Exception as e:
        logging.exception("Exception in /history/update")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/message_feedback", methods=["POST"])
async def update_message():
    await cosmos_db_ready.wait()
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for message_id
    request_json = await request.get_json()
    message_id = request_json.get("message_id", None)
    message_feedback = request_json.get("message_feedback", None)
    try:
        if not message_id:
            return jsonify({"error": "message_id is required"}), 400

        if not message_feedback:
            return jsonify({"error": "message_feedback is required"}), 400

        ## update the message in cosmos
        updated_message = await current_app.cosmos_conversation_client.update_message_feedback(
            user_id, message_id, message_feedback
        )
        if updated_message:
            return (
                jsonify(
                    {
                        "message": f"Successfully updated message with feedback {message_feedback}",
                        "message_id": message_id,
                    }
                ),
                200,
            )
        else:
            return (
                jsonify(
                    {
                        "error": f"Unable to update message {message_id}. It either does not exist or the user does not have access to it."
                    }
                ),
                404,
            )

    except Exception as e:
        logging.exception("Exception in /history/message_feedback")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/delete", methods=["DELETE"])
async def delete_conversation():
    await cosmos_db_ready.wait()
    ## get the user id from the request headers
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    try:
        if not conversation_id:
            return jsonify({"error": "conversation_id is required"}), 400

        ## make sure cosmos is configured
        if not current_app.cosmos_conversation_client:
            raise Exception("CosmosDB is not configured or not working")

        ## delete the conversation messages from cosmos first
        deleted_messages = await current_app.cosmos_conversation_client.delete_messages(
            conversation_id, user_id
        )

        ## Now delete the conversation
        deleted_conversation = await current_app.cosmos_conversation_client.delete_conversation(
            user_id, conversation_id
        )

        return (
            jsonify(
                {
                    "message": "Successfully deleted conversation and messages",
                    "conversation_id": conversation_id,
                }
            ),
            200,
        )
    except Exception as e:
        logging.exception("Exception in /history/delete")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/list", methods=["GET"])
async def list_conversations():
    await cosmos_db_ready.wait()
    offset = request.args.get("offset", 0)
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## make sure cosmos is configured
    if not current_app.cosmos_conversation_client:
        raise Exception("CosmosDB is not configured or not working")

    ## get the conversations from cosmos
    conversations = await current_app.cosmos_conversation_client.get_conversations(
        user_id, offset=offset, limit=25
    )
    if not isinstance(conversations, list):
        return jsonify({"error": f"No conversations for {user_id} were found"}), 404

    ## return the conversation ids

    return jsonify(conversations), 200


@bp.route("/history/read", methods=["POST"])
async def get_conversation():
    await cosmos_db_ready.wait()
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    if not conversation_id:
        return jsonify({"error": "conversation_id is required"}), 400

    ## make sure cosmos is configured
    if not current_app.cosmos_conversation_client:
        raise Exception("CosmosDB is not configured or not working")

    ## get the conversation object and the related messages from cosmos
    conversation = await current_app.cosmos_conversation_client.get_conversation(
        user_id, conversation_id
    )
    ## return the conversation id and the messages in the bot frontend format
    if not conversation:
        return (
            jsonify(
                {
                    "error": f"Conversation {conversation_id} was not found. It either does not exist or the logged in user does not have access to it."
                }
            ),
            404,
        )

    # get the messages for the conversation from cosmos
    conversation_messages = await current_app.cosmos_conversation_client.get_messages(
        user_id, conversation_id
    )

    ## format the messages in the bot frontend format
    messages = [
        {
            "id": msg["id"],
            "role": msg["role"],
            "content": msg["content"],
            "createdAt": msg["createdAt"],
            "feedback": msg.get("feedback"),
        }
        for msg in conversation_messages
    ]

    return jsonify({"conversation_id": conversation_id, "messages": messages}), 200


@bp.route("/history/rename", methods=["POST"])
async def rename_conversation():
    await cosmos_db_ready.wait()
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    if not conversation_id:
        return jsonify({"error": "conversation_id is required"}), 400

    ## make sure cosmos is configured
    if not current_app.cosmos_conversation_client:
        raise Exception("CosmosDB is not configured or not working")

    ## get the conversation from cosmos
    conversation = await current_app.cosmos_conversation_client.get_conversation(
        user_id, conversation_id
    )
    if not conversation:
        return (
            jsonify(
                {
                    "error": f"Conversation {conversation_id} was not found. It either does not exist or the logged in user does not have access to it."
                }
            ),
            404,
        )

    ## update the title
    title = request_json.get("title", None)
    if not title:
        return jsonify({"error": "title is required"}), 400
    conversation["title"] = title
    updated_conversation = await current_app.cosmos_conversation_client.upsert_conversation(
        conversation
    )

    return jsonify(updated_conversation), 200


@bp.route("/history/delete_all", methods=["DELETE"])
async def delete_all_conversations():
    await cosmos_db_ready.wait()
    ## get the user id from the request headers
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    # get conversations for user
    try:
        ## make sure cosmos is configured
        if not current_app.cosmos_conversation_client:
            raise Exception("CosmosDB is not configured or not working")

        conversations = await current_app.cosmos_conversation_client.get_conversations(
            user_id, offset=0, limit=None
        )
        if not conversations:
            return jsonify({"error": f"No conversations for {user_id} were found"}), 404

        # delete each conversation
        for conversation in conversations:
            ## delete the conversation messages from cosmos first
            deleted_messages = await current_app.cosmos_conversation_client.delete_messages(
                conversation["id"], user_id
            )

            ## Now delete the conversation
            deleted_conversation = await current_app.cosmos_conversation_client.delete_conversation(
                user_id, conversation["id"]
            )
        return (
            jsonify(
                {
                    "message": f"Successfully deleted conversation and messages for user {user_id}"
                }
            ),
            200,
        )

    except Exception as e:
        logging.exception("Exception in /history/delete_all")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/clear", methods=["POST"])
async def clear_messages():
    await cosmos_db_ready.wait()
    ## get the user id from the request headers
    authenticated_user = get_authenticated_user_details(request_headers=request.headers)
    user_id = authenticated_user["user_principal_id"]

    ## check request for conversation_id
    request_json = await request.get_json()
    conversation_id = request_json.get("conversation_id", None)

    try:
        if not conversation_id:
            return jsonify({"error": "conversation_id is required"}), 400

        ## make sure cosmos is configured
        if not current_app.cosmos_conversation_client:
            raise Exception("CosmosDB is not configured or not working")

        ## delete the conversation messages from cosmos
        deleted_messages = await current_app.cosmos_conversation_client.delete_messages(
            conversation_id, user_id
        )

        return (
            jsonify(
                {
                    "message": "Successfully deleted messages in conversation",
                    "conversation_id": conversation_id,
                }
            ),
            200,
        )
    except Exception as e:
        logging.exception("Exception in /history/clear_messages")
        return jsonify({"error": str(e)}), 500


@bp.route("/history/ensure", methods=["GET"])
async def ensure_cosmos():
    await cosmos_db_ready.wait()
    if not app_settings.chat_history:
        return jsonify({"error": "CosmosDB is not configured"}), 404

    try:
        success, err = await current_app.cosmos_conversation_client.ensure()
        if not current_app.cosmos_conversation_client or not success:
            if err:
                return jsonify({"error": err}), 422
            return jsonify({"error": "CosmosDB is not configured or not working"}), 500

        return jsonify({"message": "CosmosDB is configured and working"}), 200
    except Exception as e:
        logging.exception("Exception in /history/ensure")
        cosmos_exception = str(e)
        if "Invalid credentials" in cosmos_exception:
            return jsonify({"error": cosmos_exception}), 401
        elif "Invalid CosmosDB database name" in cosmos_exception:
            return (
                jsonify(
                    {
                        "error": f"{cosmos_exception} {app_settings.chat_history.database} for account {app_settings.chat_history.account}"
                    }
                ),
                422,
            )
        elif "Invalid CosmosDB container name" in cosmos_exception:
            return (
                jsonify(
                    {
                        "error": f"{cosmos_exception}: {app_settings.chat_history.conversations_container}"
                    }
                ),
                422,
            )
        else:
            return jsonify({"error": "CosmosDB is not working"}), 500


async def generate_title(conversation_messages) -> str:
    ## make sure the messages are sorted by _ts descending
    title_prompt = "Summarize the conversation so far into a 4-word or less title. Do not use any quotation marks or punctuation. Do not include any other commentary or description."

    messages = [
        {"role": msg["role"], "content": msg["content"]}
        for msg in conversation_messages
    ]
    messages.append({"role": "user", "content": title_prompt})

    try:
        azure_openai_client = await init_openai_client()
        response = await azure_openai_client.chat.completions.create(
            model=app_settings.azure_openai.model, messages=messages, temperature=1, max_tokens=64
        )

        title = response.choices[0].message.content
        return title
    except Exception as e:
        logging.exception("Exception while generating title", e)
        return messages[-2]["content"]


app = create_app()
import React, { useRef, useState, useEffect, useContext } from 'react'
import styles from './ImageFolder.module.css'
import { translateImagesApi } from '../../../api'
import { Stack } from '@fluentui/react'
import { ShieldLockRegular } from '@fluentui/react-icons'
import { getUserInfo } from '../../../api'
import { AppStateContext } from '../../../state/AppProvider'
import UploadIcon from '../../../assets/upload-icon.png'
import ImageFolderIcon from '../../../assets/imagefoldericon.png'
import { toast } from 'react-toastify'
import 'react-toastify/dist/ReactToastify.css'

const ImageTranslate = () => {
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [files, setFiles] = useState<File[]>([])
  const [language, setLanguage] = useState<string>('')
  const [isTranslating, setIsTranslating] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const appStateContext = useContext(AppStateContext)
  const [translate_tab_languages, setTranslateTabLanguages] = useState<string[] | undefined>(undefined)
  const [translate_tab_image_upload_container_text, setTranslateTabImageUploadContainerText] = useState<
    string | undefined
  >(undefined)
  const [translate_tab_image_upload_limit, setTranslateTabImageUploadLimit] = useState<number | undefined>(undefined)
  const ui = appStateContext?.state.frontendSettings?.ui
  const abortControllerRef = useRef<AbortController | null>(null)

  const AUTH_ENABLED = appStateContext?.state.frontendSettings?.auth_enabled
  const [showAuthMessage, setShowAuthMessage] = useState<boolean | undefined>()
  const prevUrlRef = useRef<string | null>(null)

  const getUserInfoList = async () => {
    if (!AUTH_ENABLED) {
      setShowAuthMessage(false)
      return
    }
    const userInfoList = await getUserInfo()
    if (userInfoList.length === 0 && window.location.hostname !== '127.0.0.1') {
      setShowAuthMessage(true)
    } else {
      setShowAuthMessage(false)
    }
  }
  useEffect(() => {
    if (!appStateContext?.state.isLoading) {
      setTranslateTabLanguages(
        ui?.translate_tab_languages || ['Chinese', 'English', 'French', 'German', 'Japanese', 'Portuguese', 'Spanish']
      )
      const limit =
        typeof ui?.translate_tab_image_upload_limit === 'number' && !isNaN(ui?.translate_tab_image_upload_limit)
          ? ui?.translate_tab_image_upload_limit
          : 50
      setTranslateTabImageUploadLimit(limit)
      setTranslateTabImageUploadContainerText(
        ui?.translate_tab_image_upload_container_text ||
          `Upload a folder of images for text translation. Only PNG, JPG, and JPEG files will be selected. Maximum of ${limit} images per folder allowed.`
      )
    }
  }, [appStateContext?.state.isLoading])

  useEffect(() => {
    if (AUTH_ENABLED !== undefined) getUserInfoList()
  }, [AUTH_ENABLED])

  const handleUploadClick = () => {
    fileInputRef.current?.click()
  }

const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
  setErrorMessage(null)
  const selectedFiles = e.target.files ? Array.from(e.target.files) : []
  const imageFiles = selectedFiles.filter(file => file.type.startsWith('image/'))
  if (imageFiles.length === 0) {
    setErrorMessage('The uploaded folder has no images. Please add at least one image file in the folder.')
    setFiles([])
    if (fileInputRef.current) fileInputRef.current.value = ''
    return
  }
  setFiles(imageFiles)
  setDownloadUrl(null)
  if (imageFiles.length > (translate_tab_image_upload_limit ?? 50)) {
    setErrorMessage(`Maximum of ${translate_tab_image_upload_limit ?? 50} images per folder allowed.`)
  } else {
    setErrorMessage(null)
  }
}

  const handleCancel = async () => {
    abortControllerRef.current?.abort()
    setFiles([])
    setDownloadUrl(null)
    setLanguage('')
    setIsTranslating(false)
    setErrorMessage(null)
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }

  const handleLanguageChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setLanguage(e.target.value)
  }

  const handleTranslate = async () => {
    if (files.length === 0 || !language) {
      setErrorMessage('Please upload image files and select a language.')
      return
    }
    setErrorMessage(null)
    setIsTranslating(true)
    abortControllerRef.current = new AbortController()
    try {
      const formData = new FormData()
      files.forEach(file => {
        formData.append('images', file)
      })
      formData.append('language', language)
      // Call new API for images translation
      const zipBlob = await translateImagesApi(formData, abortControllerRef.current.signal)
      const url = window.URL.createObjectURL(zipBlob)
      setDownloadUrl(url)
    } catch (error: any) {
      if (error.name === 'AbortError') {
        // Cancelled
      } else if (
        (error.message && error.message.includes('413')) ||
        (error.response && error.response.status === 413)
      ) {
        setErrorMessage('Maximum file size exceeded. Please insert fewer images.')
      } else {
        setErrorMessage('Error during translation. Please try again later.')
      }
    } finally {
      setIsTranslating(false)
    }
  }
  const getDownloadFileName = () => {
    if (files.length === 0 || !language) return 'translated_images.zip'
    return `translated_images_${language}.zip`
  }
  useEffect(() => {
    if (downloadUrl) {
      toast.success('Images Text Translated Successfully!', {
        position: 'top-right',
        autoClose: 3000,
        hideProgressBar: false,
        closeOnClick: true,
        pauseOnHover: true,
        draggable: true,
        progress: undefined
      })
    }
  }, [downloadUrl])

useEffect(() => {
  if (downloadUrl && prevUrlRef.current && prevUrlRef.current !== downloadUrl) {
    window.URL.revokeObjectURL(prevUrlRef.current)
  }
  prevUrlRef.current = downloadUrl
  return () => {
    if (prevUrlRef.current) {
      window.URL.revokeObjectURL(prevUrlRef.current)
    }
  }
}, [downloadUrl])
  return (
    <div className={styles.container} role="main">
      {showAuthMessage ? (
        <Stack className={styles.chatEmptyState}>
          <ShieldLockRegular
            className={styles.chatIcon}
            style={{ color: 'darkorange', height: '200px', width: '200px' }}
          />
        </Stack>
      ) : (
        <div className={styles.chatRight}>
          <div className={styles.translateCard}>
            {files.length === 0 ? (
              <div className={styles.uploadBox} onClick={handleUploadClick}>
                <input
                  type="file"
                  accept=".jpg,.jpeg,.png"
                  multiple
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  style={{ display: 'none' }}
                  // @ts-ignore
                  webkitdirectory="true"
                />
                <div className={styles.uploadIcon}>
                  <img src={UploadIcon} alt="Upload Icon" className={styles.uploadIconImage} />
                </div>
                <p>{translate_tab_image_upload_container_text}</p>
                <button className={styles.chooseButton}>Choose Folder</button>
              </div>
            ) : (
              <div className={styles.translationBox}>
                <div className={styles.pptPreview}>
                  <img src={ImageFolderIcon} alt="Images Icon" className={styles.pptIcon} />
                  <p style={{ fontSize: '1.2rem' }}>
                    {!downloadUrl
                      ? `${files.length} ${files.length === 1 ? 'image selected' : 'images selected'}`
                      : `${files.length === 1 ? 'Image Translated Successfully!' : 'Images Translated Successfully!'}`}
                  </p>
                </div>
                {!downloadUrl && (
                  <>
                    <select value={language} onChange={handleLanguageChange} className={styles.languageDropdown}>
                      <option value="">Select Language</option>
                      {translate_tab_languages?.map((lang, index) => (
                        <option key={index} value={lang}>
                          {lang}
                        </option>
                      ))}
                    </select>

                    <div className={styles.buttonGroup}>
                      <button onClick={handleCancel} className={styles.cancelButton}>
                        Cancel
                      </button>
                      <button
                        onClick={handleTranslate}
                        className={styles.translateButton}
                       disabled={
  !language ||
  isTranslating ||
  files.length === 0 ||
  files.length > (translate_tab_image_upload_limit ?? 50) ||
  !!errorMessage
}>
                        {isTranslating ? (
                          <>
                            <span>Translating...</span>
                            <span className={styles.spinner} />
                          </>
                        ) : (
                          'Translate'
                        )}
                      </button>
                    </div>
                    {errorMessage && (
                      <div style={{ color: '#c72c2c', marginTop: '0.5rem', fontSize: '1rem' }}>{errorMessage}</div>
                    )}
                  </>
                )}
              </div>
            )}

            {downloadUrl && (
              <div className={styles.buttonGroup} style={{ marginTop: '1.5rem' }}>
                <button onClick={handleCancel} className={styles.cancelButton}>
                  Start Over
                </button>
                <a href={downloadUrl} download={getDownloadFileName()} className={styles.downloadLink}>
                  <button className={styles.translateButton}>Download Zip</button>
                </a>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default ImageTranslate

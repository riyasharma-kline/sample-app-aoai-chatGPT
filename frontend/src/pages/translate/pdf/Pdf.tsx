import React, { useRef, useState, useEffect, useContext } from 'react'
import styles from './Pdf.module.css'
import { translatePdfApi } from '../../../api'
import { Stack } from '@fluentui/react'
import { ShieldLockRegular } from '@fluentui/react-icons'
import { getUserInfo } from '../../../api'
import { AppStateContext } from '../../../state/AppProvider'
import UploadIcon from '../../../assets/upload-icon.png'
import PdfIcon from '../../../assets/pdf-icon.png'
import { toast } from 'react-toastify'
import 'react-toastify/dist/ReactToastify.css'
import { Document } from 'react-pdf'
import { pdfjs } from 'react-pdf'
// Set the workerSrc property for pdfjs. This is required for pdfjs to count pdf pages correctly.
pdfjs.GlobalWorkerOptions.workerSrc = `//unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`

const Pdf = () => {
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [file, setFile] = useState<File | null>(null)
  const [language, setLanguage] = useState<string>('')
  const [isTranslating, setIsTranslating] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const appStateContext = useContext(AppStateContext)
  const [pdfPageCount, setPdfPageCount] = useState<number | null>(null)
  const [translate_tab_pdf_upload_limit, setTranslateTabPdfUploadLimit] = useState<number | undefined>(undefined)
  const [translate_tab_languages, setTranslateTabLanguages] = useState<string[] | undefined>(undefined)
  const [translate_tab_pdf_upload_container_text, setTranslateTabPdfUploadContainerText] = useState<string | undefined>(
    undefined
  )
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
      setTranslateTabPdfUploadLimit(ui?.translate_tab_pdf_upload_limit || 50)
      setTranslateTabLanguages(
        ui?.translate_tab_languages || ['Chinese', 'English', 'French', 'German', 'Japanese', 'Portuguese', 'Spanish']
      )

      setTranslateTabPdfUploadContainerText(
        ui?.translate_tab_pdf_upload_container_text || `Upload a PDF to Translate it. `
      )
    }
  }, [appStateContext?.state.isLoading])

  useEffect(() => {
    if (AUTH_ENABLED !== undefined) getUserInfoList()
  }, [AUTH_ENABLED])

  const handleUploadClick = () => {
    fileInputRef.current?.click()
  }
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      const selectedFile = e.target.files[0]
      setDownloadUrl(null)
      setPdfPageCount(null)
      if (!selectedFile.name.toLowerCase().endsWith('.pdf')) {
        setErrorMessage('Only .pdf files are supported. Please upload a .pdf file.')
        setFile(null)
        if (fileInputRef.current) fileInputRef.current.value = ''
        return
      }
      setErrorMessage(null)
      setFile(selectedFile)
    }
  }

  const handleCancel = async () => {
    abortControllerRef.current?.abort()
    setFile(null)
    setDownloadUrl(null) // reset previous translation
    setLanguage('')
    setIsTranslating(false)
    setErrorMessage(null)
    if (fileInputRef.current) {
      fileInputRef.current.value = '' // Clear the file input
    }
  }

  const handleLanguageChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setLanguage(e.target.value)
  }

  const handleTranslate = async () => {
    if (!file || !language) {
      setErrorMessage('Please upload a PDF file and select a language.')
      return
    }
    setErrorMessage(null)

    setIsTranslating(true)
    abortControllerRef.current = new AbortController()
    try {
      const translatedBlob = await translatePdfApi(file, language, abortControllerRef.current.signal)
      const url = window.URL.createObjectURL(translatedBlob)
      setDownloadUrl(url)
    } catch (error: any) {
      if (error.name === 'AbortError') {
      } else if (
        (error.message && error.message.includes('413')) ||
        (error.response && error.response.status === 413)
      ) {
        setErrorMessage('Maximum file size exceeded. Please insert a shorter file.')
      } else {
        setErrorMessage('Error during translation. Please try using a shorter file or try again later.')
      }
    } finally {
      setIsTranslating(false)
    }
  }
  const getDownloadFileName = () => {
    if (!file || !language) return 'translated_document.pdf'
    const name = file.name.replace(/\.[^/.]+$/, '') // remove extension
    return `${name}_${language}.pdf`
  }
  useEffect(() => {
    if (downloadUrl) {
      toast.success('PDF Translated Successfully!', {
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

  // Clean up previous object URL to avoid memory leaks
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
            {!file ? (
              <div className={styles.uploadBox} onClick={handleUploadClick}>
                <input
                  type="file"
                  accept=".pdf"
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  style={{ display: 'none' }}
                />
                <div className={styles.uploadIcon}>
                  <img src={UploadIcon} alt="Upload Icon" className={styles.uploadIconImage} />
                </div>
                <p>{translate_tab_pdf_upload_container_text}</p>
                <button className={styles.chooseButton}>Choose file</button>
              </div>
            ) : (
              <div className={styles.translationBox}>
                <div className={styles.pptPreview}>
                  <img src={PdfIcon} alt="PDF Icon" className={styles.pptIcon} />
                  <p style={{ fontSize: '1.5rem' }}>{!downloadUrl ? file.name : 'PDF Translated Successfully!'}</p>
                  {file && (
                    <Document
                      file={file}
                      onLoadSuccess={({ numPages }: { numPages: number }) => {
                        setPdfPageCount(numPages)
                        if (translate_tab_pdf_upload_limit !== undefined) {
                          if (numPages > translate_tab_pdf_upload_limit) {
                            setErrorMessage(
                              `Please upload a PDF file with a maximum of ${translate_tab_pdf_upload_limit} pages.`
                            )
                          } else if (numPages === 0) {
                            setErrorMessage('The uploaded file contains no pages. Please upload a valid PDF file.')
                          } else {
                            setErrorMessage(null)
                          }
                        } else {
                          setErrorMessage(null)
                        }
                      }}
                      onLoadError={(error: any) => {
                        setErrorMessage('Could not read PDF file. Please try another file.')
                        setPdfPageCount(null)
                      }}
                    />
                  )}
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
                          (pdfPageCount !== null && pdfPageCount === 0) ||
                          (translate_tab_pdf_upload_limit !== undefined &&
                            pdfPageCount !== null &&
                            pdfPageCount > translate_tab_pdf_upload_limit)
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
                  <button className={styles.translateButton}>Download</button>
                </a>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default Pdf

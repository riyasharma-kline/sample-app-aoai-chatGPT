import React, { useRef, useState, useEffect, useContext } from 'react'
import styles from './Translate.module.css'
import { translateApi } from '../../api'
import { Stack } from '@fluentui/react'
import { ShieldLockRegular } from '@fluentui/react-icons'
import { getUserInfo } from '../../api'
import { AppStateContext } from '../../state/AppProvider'
import Contoso from '../../assets/Contoso.svg'
import UploadIcon from '../../assets/upload-icon.png'
import PPTIcon from '../../assets/ppt-icon.png'
import { ToastContainer, toast } from 'react-toastify'
import 'react-toastify/dist/ReactToastify.css'
import JSZip from 'jszip'

const Translate = () => {
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [file, setFile] = useState<File | null>(null)
  const [language, setLanguage] = useState<string>('')
  const [isTranslating, setIsTranslating] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const appStateContext = useContext(AppStateContext)
  const [translate_tab_description_line1, setTranslateTabDescriptionLine1] = useState<string | undefined>(undefined)
  const [translate_tab_description_line2, setTranslateTabDescriptionLine2] = useState<string | undefined>(undefined)
  const [translate_tab_title, setTranslateTabTitle] = useState<string | undefined>(undefined)
  const [translate_tab_slide_limit_enable, setTranslateTabSlideLimitEnable] = useState<boolean | undefined>(undefined)
  const [translate_tab_slide_limit, setTranslateTabSlideLimit] = useState<number | undefined>(undefined)
  const [translate_tab_languages, setTranslateTabLanguages] = useState<string[] | undefined>(undefined)
  const [logo, setLogo] = useState('')
  const ui = appStateContext?.state.frontendSettings?.ui
  const abortControllerRef = useRef<AbortController | null>(null)

  const AUTH_ENABLED = appStateContext?.state.frontendSettings?.auth_enabled
  const [showAuthMessage, setShowAuthMessage] = useState<boolean | undefined>()
  const [slideCount, setSlideCount] = useState<number | null>(null)

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
      setLogo(ui?.chat_logo || ui?.logo || Contoso)
      setTranslateTabDescriptionLine1(ui?.translate_tab_description_line1 || 'Easily translate the reports in any language with the power of AI.')
      setTranslateTabDescriptionLine2(ui?.translate_tab_description_line2 || 'After translation, please check the downloaded file for font, layout and translation errors')
      setTranslateTabTitle(ui?.translate_tab_title || 'Translate it the Kline way') 
      setTranslateTabSlideLimitEnable(ui?.translate_tab_slide_limit_enable)
      setTranslateTabSlideLimit(ui?.translate_tab_slide_limit || 50)
      setTranslateTabLanguages(ui?.translate_tab_languages || ['Chinese','English','French', 'German', 'Japanese', 'Portuguese', 'Spanish'])
    }
  }, [appStateContext?.state.isLoading])

  useEffect(() => {
    if (AUTH_ENABLED !== undefined) getUserInfoList()
  }, [AUTH_ENABLED])

  const handleUploadClick = () => {
    fileInputRef.current?.click()
  }

  const countSlidesInPptx = async (file: File): Promise<number> => {
    const arrayBuffer = await file.arrayBuffer()
    const zip = await JSZip.loadAsync(arrayBuffer)
    const slideFiles = Object.keys(zip.files).filter((name) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(name)
    )
    return slideFiles.length
  }

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      const selectedFile = e.target.files[0]
      setDownloadUrl(null) // reset previous translation
      let count: number | null = null
      if (!selectedFile.name.endsWith('.pptx')) {
        setErrorMessage('Only .pptx files are supported. Please upload a .pptx file.')
        setFile(null)
        setSlideCount(null)
        if (fileInputRef.current) fileInputRef.current.value = ''
        return
      }
      try {
        count = await countSlidesInPptx(selectedFile)
        setSlideCount(count)
        if (translate_tab_slide_limit_enable && translate_tab_slide_limit !== undefined) {
          if (count > translate_tab_slide_limit) {
            setErrorMessage(`Please upload a PowerPoint file with a maximum of ${translate_tab_slide_limit} slides.`)
          } else {
            setErrorMessage(null)
          }
        } else {
          setErrorMessage(null)
        }
      } catch (err) {
        setErrorMessage('Could not read PPTX file. Please try another file.')
        setFile(null)
        setSlideCount(null)
        if (fileInputRef.current) fileInputRef.current.value = ''
        return
      }
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
      setErrorMessage('Please upload a PPT file and select a language.')
      return
    }
    setErrorMessage(null)

    const formData = new FormData()
    formData.append('file', file)
    formData.append('language', language)

    setIsTranslating(true)
    abortControllerRef.current = new AbortController()
    try {
      const translatedBlob = await translateApi(file, language, abortControllerRef.current.signal)

      const url = window.URL.createObjectURL(translatedBlob)

      setDownloadUrl(url)
    } catch (error: any) {
      if (error.name === 'AbortError') {
        // Cancelled, do nothing or show a message
      } else if (
        (error.message && error.message.includes('413')) ||
        (error.response && error.response.status === 413)
      ) {
        setErrorMessage('Maximum file size exceeded. Please insert a shorter file.')
      } else {
        setErrorMessage('Error during translation. Please try again later.')
      }
    } finally {
      setIsTranslating(false)
    }
  }
  const getDownloadFileName = () => {
    if (!file || !language) return 'translated_presentation.pptx'
    const name = file.name.replace(/\.[^/.]+$/, '') // remove extension
    return `${name}_${language}.pptx`
  }
  useEffect(() => {
    if (downloadUrl) {
      toast.success('PPT Translated Successfully!', {
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
  return (
    <div className={styles.container} role="main">
      {showAuthMessage ? (
        <Stack className={styles.chatEmptyState}>
          <ShieldLockRegular
            className={styles.chatIcon}
            style={{ color: 'darkorange', height: '200px', width: '200px' }}
          />
          <h1 className={styles.chatEmptyStateTitle}>Authentication Not Configured</h1>
          <h2 className={styles.chatEmptyStateSubtitle}>
            This app does not have authentication configured. Please add an identity provider by finding your app in the{' '}
            <a href="https://portal.azure.com/" target="_blank">
              Azure Portal
            </a>
            and following{' '}
            <a
              href="https://learn.microsoft.com/en-us/azure/app-service/scenario-secure-app-authentication-app-service#3-configure-authentication-and-authorization"
              target="_blank">
              these instructions
            </a>
            .
          </h2>
          <h2 className={styles.chatEmptyStateSubtitle} style={{ fontSize: '20px' }}>
            <strong>Authentication configuration takes a few minutes to apply. </strong>
          </h2>
          <h2 className={styles.chatEmptyStateSubtitle} style={{ fontSize: '20px' }}>
            <strong>If you deployed in the last 10 minutes, please wait and reload the page after 10 minutes.</strong>
          </h2>
        </Stack>
      ) : (
        <div className={styles.chatLayout} style={{ padding: downloadUrl ? '6rem' : '4rem' }}>
          <div className={styles.chatLeft}>
            <Stack className={styles.chatEmptyState}>
              <img src={logo} className={styles.chatIcon} aria-hidden="true" />
              <h1 className={styles.chatEmptyStateTitle}>{translate_tab_title}</h1>
              <h2 className={styles.chatEmptyStateSubtitle}>
                {translate_tab_description_line1}
              </h2>
              <h2 className={styles.chatEmptyStateSubtitle} style={{ marginTop: '0px' }}>
                {translate_tab_description_line2}
              </h2>
            </Stack>
          </div>
          <div className={styles.chatRight}>
            <div className={styles.translateCard}>
              {!file ? (
                <div className={styles.uploadBox} onClick={handleUploadClick}>
                  <input
                    type="file"
                    accept=".pptx"
                    ref={fileInputRef}
                    onChange={handleFileChange}
                    style={{ display: 'none' }}
                  />
                  <div className={styles.uploadIcon}>
                    <img src={UploadIcon} alt="Upload Icon" className={styles.uploadIconImage} />
                  </div>
                  <p>Drag and Drop a PowerPoint to Translate it.</p>
                  <button className={styles.chooseButton}>Choose file</button>
                </div>
              ) : (
                <div className={styles.translationBox}>
                  <div className={styles.pptPreview}>
                    <img src={PPTIcon} alt="PPT Icon" className={styles.pptIcon} />
                    <p style={{ fontSize: '1.5rem' }}>{!downloadUrl ? file.name : 'PPT Translated Successfully!'}</p>
                  </div>
                  {!downloadUrl && (
                    <>
                      <select value={language} onChange={handleLanguageChange} className={styles.languageDropdown} >
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
                          disabled={!language || isTranslating || (translate_tab_slide_limit_enable && translate_tab_slide_limit !== undefined && slideCount !== null && slideCount > translate_tab_slide_limit)}>
                          {isTranslating ? (
                            <>
                              <span>Translating...</span>
                              <span
                                style={{
                                  marginLeft: 8,
                                  display: 'inline-block',
                                  width: 16,
                                  height: 16,
                                  border: '2px solid #ccc',
                                  borderTop: '2px solid #333',
                                  borderRadius: '50%',
                                  animation: 'spin 1s linear infinite',
                                  verticalAlign: 'middle'
                                }}
                              />
                            </>
                          ) : (
                            'Translate'
                          )}
                          <style>
                            {`
          @keyframes spin {
            0% { transform: rotate(0deg);}
            100% { transform: rotate(360deg);}
          }
        `}
                          </style>
                        </button>
                      </div>
                      {errorMessage && (
                        <div style={{ color: '#c72c2c', marginTop: '0.5rem', fontSize: '1rem' }}>{errorMessage}</div>
                      )}
                    </>
                  )}
                </div>
              )}
              <ToastContainer />

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
        </div>
      )}
    </div>
  )
}

export default Translate

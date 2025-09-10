import { useState, useEffect, useContext, useRef } from 'react'
import styles from './Tidy.module.css'
import Contoso from '../../assets/Contoso.svg'
import { tidyApi } from '../../api'
import { Stack } from '@fluentui/react'
import { ShieldLockRegular } from '@fluentui/react-icons'
import { getUserInfo } from '../../api'
import { AppStateContext } from '../../state/AppProvider'
import UploadIcon from '../../assets/upload-icon.png'
import PPTIcon from '../../assets/ppt-icon.png'
import { toast } from 'react-toastify'
import 'react-toastify/dist/ReactToastify.css'
import JSZip from 'jszip'

const Tidy = () => {
  const appStateContext = useContext(AppStateContext)
  const [tidy_tab_description_line1, setTidyTabDescriptionLine1] = useState<string | undefined>(undefined)
  const [tidy_tab_description_line2, setTidyTabDescriptionLine2] = useState<string | undefined>(undefined)
  const [tidy_tab_title, setTidyTabTitle] = useState<string | undefined>(undefined)
  const [tidy_tab_slide_limit, setTidyTabSlideLimit] = useState<number | undefined>(undefined)
  const [tidy_tab_slide_upload_container_text, setTidyTabSlideUploadContainerText] = useState<string | undefined>(
    undefined
  )

  const [logo, setLogo] = useState('')
  const ui = appStateContext?.state.frontendSettings?.ui
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const [file, setFile] = useState<File | null>(null)
  const [isTranslating, setIsTranslating] = useState(false)
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)
  const abortControllerRef = useRef<AbortController | null>(null)

  const AUTH_ENABLED = appStateContext?.state.frontendSettings?.auth_enabled
  const [showAuthMessage, setShowAuthMessage] = useState<boolean | undefined>()
  const [slideCount, setSlideCount] = useState<number | null>(null)
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
      setLogo(ui?.chat_logo || ui?.logo || Contoso)
      setTidyTabDescriptionLine1(
        ui?.tidy_tab_description_line1 || 'Easily tidy up the reports in any language with the power of AI.'
      )
      setTidyTabDescriptionLine2(
        ui?.tidy_tab_description_line2 || 'After tidying, please check the downloaded file for font and layout errors'
      )
      setTidyTabTitle(ui?.tidy_tab_title || 'Tidy it the Kline way')
      setTidyTabSlideLimit(ui?.tidy_tab_slide_limit || 50)
      setTidyTabSlideUploadContainerText(ui?.tidy_tab_slide_upload_container_text || 'Upload a PowerPoint to Tidy it.')
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
    const slideFiles = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
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
        if (tidy_tab_slide_limit !== undefined) {
          if (count > tidy_tab_slide_limit) {
            setErrorMessage(`Please upload a PowerPoint file with a maximum of ${tidy_tab_slide_limit} slides.`)
          } else if (count === 0) {
            setErrorMessage('The uploaded file contains no slides. Please upload a valid PowerPoint file.')
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
    setIsTranslating(false)
    setErrorMessage(null)
    if (fileInputRef.current) {
      fileInputRef.current.value = '' // Clear the file input
    }
  }

  const handleTranslate = async () => {
    if (!file) {
      setErrorMessage('Please upload a PPT file.')
      return
    }
    setErrorMessage(null)

    const formData = new FormData()
    formData.append('file', file)

    setIsTranslating(true)
    abortControllerRef.current = new AbortController()
    try {
      const tidiedBlob = await tidyApi(file, abortControllerRef.current.signal)

      const url = window.URL.createObjectURL(tidiedBlob)

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
    if (!file) return 'tidied_presentation.pptx'
    const name = file.name.replace(/\.[^/.]+$/, '') // remove extension
    return `${name}_tidied.pptx`
  }
  useEffect(() => {
    if (downloadUrl) {
      toast.success('PPT Tidied Successfully!', {
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
        <div className={styles.chatLayout}>
          <div className={styles.chatLeft}>
            <Stack className={styles.chatEmptyState}>
              <img src={logo} className={styles.chatIcon} aria-hidden="true" />
              <h1 className={styles.chatEmptyStateTitle}>{tidy_tab_title}</h1>
              <h2 className={styles.chatEmptyStateSubtitle}>{tidy_tab_description_line1}</h2>
              <h2 className={styles.chatEmptyStateSubtitle} style={{ marginTop: '0px' }}>
                {tidy_tab_description_line2}
              </h2>
            </Stack>
          </div>
          <div className={styles.chatRight}>
            {' '}
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
                  <div className={styles.tidyCard}>
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
                        <p>{tidy_tab_slide_upload_container_text}</p>
                        <button className={styles.chooseButton}>Choose file</button>
                      </div>
                    ) : (
                      <div className={styles.translationBox}>
                        <div className={styles.pptPreview}>
                          <img src={PPTIcon} alt="PPT Icon" className={styles.pptIcon} />
                          <p style={{ fontSize: '1.5rem' }}>{!downloadUrl ? file.name : 'PPT Tidied Successfully!'}</p>
                        </div>
                        {!downloadUrl && (
                          <>
                            <div className={styles.buttonGroup}>
                              <button onClick={handleCancel} className={styles.cancelButton}>
                                Cancel
                              </button>
                              <button
                                onClick={handleTranslate}
                                className={styles.tidyButton}
                                 disabled={
    
    isTranslating ||
    (slideCount !== null && slideCount === 0) ||
    (tidy_tab_slide_limit !== undefined &&
      slideCount !== null &&
      slideCount > tidy_tab_slide_limit)
  }>
                                {isTranslating ? (
                                  <>
                                    <span>Tidying...</span>
                                    <span className={styles.spinner} />
                                  </>
                                ) : (
                                  'Tidy'
                                )}
                              </button>
                            </div>
                            {errorMessage && (
                              <div style={{ color: '#c72c2c', marginTop: '0.5rem', fontSize: '1rem' }}>
                                {errorMessage}
                              </div>
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
                          <button className={styles.tidyButton}>Download</button>
                        </a>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default Tidy

import { useState, useEffect, useContext, useRef } from 'react'
import styles from './Translate.module.css'
import { Stack } from '@fluentui/react'
import { ShieldLockRegular } from '@fluentui/react-icons'
import { getUserInfo } from '../../api'
import { AppStateContext } from '../../state/AppProvider'
import Contoso from '../../assets/Contoso.svg'
import PPTIcon from '../../assets/ppt-icon.png'
import PDFIcon from '../../assets/pdf-icon.png'
import ImageFolderIcon from '../../assets/imagefoldericon.png'
import Powerpoint from './powerpoint/Powerpoint'
import ImageFolder from './imageFolder/ImageFolder'
import Pdf from './pdf/Pdf'
import 'react-toastify/dist/ReactToastify.css'

const Translate = () => {
  const [translateType, setTranslateType] = useState<'powerpoint' | 'image folder' | 'pdf'>('powerpoint')
  const [dropdownOpen, setDropdownOpen] = useState(false)
  const appStateContext = useContext(AppStateContext)
  const [translate_tab_description_line1, setTranslateTabDescriptionLine1] = useState<string | undefined>(undefined)
  const [translate_tab_description_line2, setTranslateTabDescriptionLine2] = useState<string | undefined>(undefined)
  const [translate_tab_title, setTranslateTabTitle] = useState<string | undefined>(undefined)
  const [logo, setLogo] = useState('')
  const ui = appStateContext?.state.frontendSettings?.ui
  const dropdownRef = useRef<HTMLDivElement>(null)
  const AUTH_ENABLED = appStateContext?.state.frontendSettings?.auth_enabled
  const [showAuthMessage, setShowAuthMessage] = useState<boolean | undefined>()

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
      setTranslateTabDescriptionLine1(
        ui?.translate_tab_description_line1 || 'Easily translate the reports in any language with the power of AI.'
      )
      setTranslateTabDescriptionLine2(
        ui?.translate_tab_description_line2 ||
          'After translation, please check the downloaded file for font, layout and translation errors'
      )
      setTranslateTabTitle(ui?.translate_tab_title || 'Translate it the Kline way')
    }
  }, [appStateContext?.state.isLoading])

  useEffect(() => {
    if (AUTH_ENABLED !== undefined) getUserInfoList()
  }, [AUTH_ENABLED])

  useEffect(() => {
    if (!dropdownOpen) return
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setDropdownOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [dropdownOpen])
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
              <div className={styles.dropdownContainer}>
                <div className={styles.dropdownWrapper}>
                  <div className={styles.dropdownRelative} ref={dropdownRef}>
                    <div className={styles.dropdownButton} onClick={() => setDropdownOpen(v => !v)}>
                      {translateType === 'powerpoint' && (
                        <img src={PPTIcon} alt="PowerPoint" className={styles.dropdownIcon} />
                      )}
                      {translateType === 'image folder' && (
                        <img src={ImageFolderIcon} alt="Image Folder" className={styles.dropdownIcon} />
                      )}
                      {translateType === 'pdf' && <img src={PDFIcon} alt="PDF" className={styles.dropdownIcon} />}
                      <span className={styles.dropdownLabel}>
                        {translateType === 'powerpoint' && 'PowerPoint'}
                        {translateType === 'image folder' && 'Image Folder'}
                        {translateType === 'pdf' && 'PDF'}
                      </span>
                      <span className={styles.dropdownArrow} aria-hidden="true">
                        <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                          <path
                            d="M4 6l4 4 4-4"
                            stroke="#222"
                            strokeWidth="2"
                            strokeLinecap="round"
                            strokeLinejoin="round"
                          />
                        </svg>
                      </span>
                    </div>
                    {dropdownOpen && (
                      <div className={styles.dropdownMenu}>
                        <div
                          className={styles.dropdownMenuItem}
                          onClick={() => {
                            setTranslateType('powerpoint')
                            setDropdownOpen(false)
                          }}
                          onMouseOver={e => e.currentTarget.classList.add(styles.dropdownMenuItemHover)}
                          onMouseOut={e => e.currentTarget.classList.remove(styles.dropdownMenuItemHover)}>
                          <img src={PPTIcon} alt="PowerPoint" className={styles.dropdownIcon} />
                          <span className={styles.dropdownLabel}>PowerPoint</span>
                        </div>
                        <div
                          className={styles.dropdownMenuItem}
                          onClick={() => {
                            setTranslateType('image folder')
                            setDropdownOpen(false)
                          }}
                          onMouseOver={e => e.currentTarget.classList.add(styles.dropdownMenuItemHover)}
                          onMouseOut={e => e.currentTarget.classList.remove(styles.dropdownMenuItemHover)}>
                          <img src={ImageFolderIcon} alt="Image Folder" className={styles.dropdownIcon} />
                          <span className={styles.dropdownLabel}>Image Folder</span>
                        </div>
                        <div
                          className={styles.dropdownMenuItem}
                          onClick={() => {
                            setTranslateType('pdf')
                            setDropdownOpen(false)
                          }}
                          onMouseOver={e => e.currentTarget.classList.add(styles.dropdownMenuItemHover)}
                          onMouseOut={e => e.currentTarget.classList.remove(styles.dropdownMenuItemHover)}>
                          <img src={PDFIcon} alt="PDF" className={styles.dropdownIcon} />
                          <span className={styles.dropdownLabel}>PDF</span>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <h1 className={styles.chatEmptyStateTitle}>{translate_tab_title}</h1>
              <h2 className={styles.chatEmptyStateSubtitle}>{translate_tab_description_line1}</h2>
              <h2 className={styles.chatEmptyStateSubtitle} style={{ marginTop: '0px' }}>
                {translate_tab_description_line2}
              </h2>
            </Stack>
          </div>
          <div className={styles.chatRight}>
            {translateType === 'powerpoint' && <Powerpoint />}
            {translateType === 'image folder' && <ImageFolder />}
            {translateType === 'pdf' && <Pdf />}
          </div>
        </div>
      )}
    </div>
  )
}

export default Translate

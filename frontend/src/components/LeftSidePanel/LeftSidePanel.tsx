import React, { useState, useEffect, useContext } from 'react'
import { Dropdown, IDropdownOption, Stack, Text, IconButton } from '@fluentui/react'
import { AppStateContext } from '../../state/AppProvider'
import styles from './LeftSidePanel.module.css'

interface LeftSidePanelProps {
  onClose: () => void
  onWritingStyleChange: (style: string) => void
  selectedWritingStyle?: string
}

export function LeftSidePanel({ onClose, onWritingStyleChange, selectedWritingStyle }: LeftSidePanelProps) {
  const appStateContext = useContext(AppStateContext)

  const [writingStyles, setWritingStyles] = useState<IDropdownOption[]>([])
  const [leftSidePanelHeading, setLeftSidePanelHeading] = useState<string>('Settings')
  const [selectedStyle, setSelectedStyle] = useState<IDropdownOption | undefined>(undefined)

  useEffect(() => {
    const stylesFromConfig = appStateContext?.state.frontendSettings?.ui?.writing_style_options
    const leftSidePanelHeadingFromConfig = appStateContext?.state.frontendSettings?.ui?.left_side_panel_heading

    if (leftSidePanelHeadingFromConfig) {
      setLeftSidePanelHeading(leftSidePanelHeadingFromConfig)
    }

    if (!stylesFromConfig) return

    const options = stylesFromConfig.map((style: string) => ({
      key: style,
      text: style
    }))

    setWritingStyles(options)

    if (!selectedWritingStyle && options.length > 0) {
      setSelectedStyle(options[0])
      onWritingStyleChange(options[0].key as string)
    }
  }, [appStateContext])
  useEffect(() => {
    if (!selectedWritingStyle || writingStyles.length === 0) return

    const found = writingStyles.find(o => o.key === selectedWritingStyle)
    setSelectedStyle(found)
  }, [selectedWritingStyle, writingStyles])

  const onStyleChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setSelectedStyle(option)
    if (option) {
      onWritingStyleChange(option.key as string)
    }
  }

  return (
    <aside className={styles.sidebar} aria-label="Settings panel">
      {/* Header */}
      <div className={styles.header}>
        <Text variant="xLarge" className={styles.title}>
          {leftSidePanelHeading}
        </Text>
        <IconButton iconProps={{ iconName: 'Cancel' }} ariaLabel="Close panel" onClick={onClose} />
      </div>

      {/* Content */}
      <Stack tokens={{ childrenGap: 12 }} className={styles.content}>
        <Text variant="mediumPlus" className={styles.sectionLabel}>
          Writing Style
        </Text>
        <Dropdown
          placeholder="Select a writing style"
          options={writingStyles}
          selectedKey={selectedStyle?.key}
          onChange={onStyleChange}
          styles={{ dropdown: { width: '100%' } }}
        />
      </Stack>
    </aside>
  )
}

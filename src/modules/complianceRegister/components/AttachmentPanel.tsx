import * as React from 'react';
import styles from './AttachmentPanel.module.scss';
import { IAttachment } from '../models/IComplianceItem';

interface IAttachmentPanelProps {
  attachments: IAttachment[];
  siteUrl: string;
  loading: boolean;
  onUpload: (file: File) => Promise<void>;
  onRemove: (fileName: string) => Promise<void>;
}

const AttachmentPanel: React.FC<IAttachmentPanelProps> = ({ attachments, siteUrl, loading, onUpload, onRemove }) => {
  const [selectedFile, setSelectedFile] = React.useState<string | null>(null);
  const [uploading, setUploading] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleUploadClick = (): void => {
    if (fileInputRef.current) fileInputRef.current.click();
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      await onUpload(file);
    } finally {
      setUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleRemove = async (): Promise<void> => {
    if (!selectedFile) return;
    await onRemove(selectedFile);
    setSelectedFile(null);
  };

  const handleDownload = (): void => {
    if (!selectedFile) return;
    const att = attachments.find(a => a.FileName === selectedFile);
    if (att) window.open(`${siteUrl}${att.ServerRelativeUrl}`, '_blank');
  };

  return (
    <div className={styles.attachmentPanel}>
      <div className={styles.sectionHeader}>
        <span className={styles.chevron}>›</span>
        <span className={styles.sectionTitle}>ATTACHMENT</span>
      </div>
      <div className={styles.attachmentContent}>
        {/* Actions bar */}
        <div className={styles.attachActions}>
          <span className={styles.attachCount}>{attachments.length}</span>
          <button className={styles.attachBtn} onClick={handleUploadClick} disabled={uploading} id="btn-upload-attachment">
            {uploading ? '↑ Uploading…' : '⬆ Upload'}
          </button>
          <button className={styles.attachBtnSecondary} onClick={handleRemove} disabled={!selectedFile} id="btn-remove-attachment">
            ✖ Remove
          </button>
          <button className={styles.attachBtnSecondary} onClick={handleDownload} disabled={!selectedFile} id="btn-download-attachment">
            ⬇ Download
          </button>
          <input ref={fileInputRef} type="file" style={{ display: 'none' }} onChange={handleFileChange} id="file-input-hidden" />
        </div>

        {/* File list */}
        {loading && <div className={styles.loadingText}>Loading attachments…</div>}
        {!loading && attachments.length === 0 && (
          <div className={styles.emptyAttach}>No attachments</div>
        )}
        <ul className={styles.fileList}>
          {attachments.map(att => (
            <li
              key={att.FileName}
              className={`${styles.fileItem} ${selectedFile === att.FileName ? styles.selectedFile : ''}`}
              onClick={() => setSelectedFile(att.FileName === selectedFile ? null : att.FileName)}
              id={`attach-${att.FileName.replace(/[^a-zA-Z0-9]/g, '-')}`}
            >
              <span className={styles.fileIcon}>📎</span>
              <span className={styles.fileName}>{att.FileName}</span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
};

export default AttachmentPanel;

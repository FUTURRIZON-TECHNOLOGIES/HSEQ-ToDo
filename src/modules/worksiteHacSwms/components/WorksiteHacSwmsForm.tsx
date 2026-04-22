import * as React from 'react';
import {
    DatePicker,
    DefaultButton,
    Dropdown,
    IDropdownOption,
    Icon,
    PrimaryButton,
    TextField
} from '@fluentui/react';
import { IWorksiteHacSwmsRecord } from '../models/IWorksiteHacSwmsRecord';
import { projectOptions, statusOptions } from '../data/mockWorksiteHacSwmsData';
import HacPdfPreview from './HacPdfPreview';
import styles from '../WorksiteHacSwmsModule.module.scss';

export interface IWorksiteHacSwmsFormProps {
    item: IWorksiteHacSwmsRecord;
    onClose: () => void;
}

const toOptions = (values: string[]): IDropdownOption[] => values.map(value => ({ key: value, text: value }));

const WorksiteHacSwmsForm: React.FC<IWorksiteHacSwmsFormProps> = ({ item, onClose }) => {
    const [formData, setFormData] = React.useState<IWorksiteHacSwmsRecord>(item);

    React.useEffect(() => {
        setFormData(item);
    }, [item]);

    return (
        <div className={styles.formShell}>
            <div className={styles.detailToolbar}>
                <div className={styles.detailTitle}>
                    <Icon iconName="ClipboardList" />
                    <span>Worksite HAC & SWMS: {formData.number}</span>
                </div>
                <div className={styles.detailActions}>
                    <PrimaryButton iconProps={{ iconName: 'Save' }} text="Save" />
                    <DefaultButton iconProps={{ iconName: 'SaveAndClose' }} text="Save & Close" onClick={onClose} />
                    <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" />
                    <DefaultButton iconProps={{ iconName: 'Cancel' }} text="Close" onClick={onClose} />
                </div>
            </div>

            <div className={styles.formContent}>
                <div className={styles.editorColumn}>
                    <div className={styles.formGrid}>
                        <TextField label="Number" value={formData.number} disabled />
                        <DatePicker
                            label="Date"
                            value={new Date(formData.date)}
                            onSelectDate={date => setFormData({ ...formData, date: date?.toISOString() || formData.date })}
                        />
                        <Dropdown
                            label="Project"
                            selectedKey={formData.project}
                            options={toOptions(projectOptions)}
                            onChange={(_, option) => setFormData({ ...formData, project: String(option?.key || '') })}
                        />
                        <TextField
                            label="Scope of Works"
                            multiline
                            rows={3}
                            value={formData.scopeOfWorks}
                            onChange={(_, value) => setFormData({ ...formData, scopeOfWorks: value || '' })}
                        />
                        <TextField
                            label="Work Addresses"
                            multiline
                            rows={3}
                            value={formData.workAddresses}
                            onChange={(_, value) => setFormData({ ...formData, workAddresses: value || '' })}
                        />
                        <TextField
                            label="Supervisor/Coordinator"
                            value={formData.supervisor}
                            onChange={(_, value) => setFormData({ ...formData, supervisor: value || '' })}
                        />
                        <TextField label="Geo Location" value={formData.geoLocation} />
                        <Dropdown
                            label="Status"
                            selectedKey={formData.status}
                            options={toOptions(statusOptions)}
                            onChange={(_, option) => setFormData({ ...formData, status: String(option?.key || '') })}
                            className={styles.statusDropdown}
                        />
                    </div>

                    <PrimaryButton className={styles.continueButton} iconProps={{ iconName: 'Edit' }} text="Continue" />

                    <div className={styles.sectionBlock}>
                        <div className={styles.sectionTitle}>SWMS Used</div>
                        <div className={styles.swmsPicker}>
                            <Dropdown
                                placeholder="Search SWMS..."
                                options={[
                                    { key: '015', text: '015 - IAC WHS Training Facility Safety' },
                                    { key: '018', text: '018 - Electrical Training Controls' }
                                ]}
                            />
                            <DefaultButton text="Add to List" disabled />
                        </div>
                        <div className={styles.swmsList}>
                            {formData.swmsUsed.map(swms => (
                                <div className={styles.swmsRow} key={swms.id}>
                                    <button type="button" title="Remove"><Icon iconName="Cancel" /></button>
                                    <a>{swms.id}</a>
                                    <span>{swms.title}</span>
                                </div>
                            ))}
                        </div>
                    </div>

                    <div className={styles.sectionBlock}>
                        <div className={styles.sectionTitle}>Photos & Attachments</div>
                        <div className={styles.uploadDrop}>
                            <span>Drop file(s) here</span>
                            <DefaultButton text="Upload" />
                        </div>
                    </div>

                    <div className={styles.timelineBlock}>
                        <div className={styles.timelineHeader}>
                            <Icon iconName="CircleAddition" />
                            <span>New Timeline</span>
                        </div>
                        {formData.timeline.map(entry => (
                            <div className={styles.timelineEntry} key={entry.id}>
                                <div className={styles.avatar}>{entry.avatarText}</div>
                                <div>
                                    <strong>{entry.name}</strong>
                                    <span>{entry.dateText}</span>
                                    <p>{entry.description}</p>
                                </div>
                            </div>
                        ))}
                    </div>
                </div>

                <div className={styles.previewColumn}>
                    <HacPdfPreview item={formData} />
                </div>
            </div>
        </div>
    );
};

export default WorksiteHacSwmsForm;

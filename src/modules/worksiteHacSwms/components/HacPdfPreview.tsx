import * as React from 'react';
import { Icon } from '@fluentui/react';
import { IWorksiteHacSwmsRecord } from '../models/IWorksiteHacSwmsRecord';
import styles from '../WorksiteHacSwmsModule.module.scss';

export interface IHacPdfPreviewProps {
    item: IWorksiteHacSwmsRecord;
}

const formatLongDate = (dateValue: string): string => {
    const date = new Date(dateValue);
    return date.toLocaleDateString('en-AU', {
        weekday: 'long',
        day: '2-digit',
        month: 'long',
        year: 'numeric'
    });
};

const HacPdfPreview: React.FC<IHacPdfPreviewProps> = ({ item }) => {
    return (
        <div className={styles.pdfViewer}>
            <div className={styles.pdfToolbar}>
                <div className={styles.pdfFile}>
                    <Icon iconName="GlobalNavButton" />
                    <span>HAC_{item.number}.pdf</span>
                </div>
                <div className={styles.pdfTools}>
                    <span className={styles.pdfPageIndicator}>1 / 2</span>
                    <span className={styles.pdfDivider} />
                    <button type="button" title="Zoom out"><Icon iconName="ZoomOut" /></button>
                    <span className={styles.zoomValue}>100%</span>
                    <button type="button" title="Zoom in"><Icon iconName="ZoomIn" /></button>
                    <span className={styles.pdfDivider} />
                    <button type="button" title="Download"><Icon iconName="Download" /></button>
                    <button type="button" title="Print"><Icon iconName="Print" /></button>
                    <button type="button" title="More"><Icon iconName="MoreVertical" /></button>
                </div>
            </div>

            <div className={styles.pdfStage}>
                <div className={styles.pdfPage}>
                    <div className={styles.pdfHeader}>
                        <div className={styles.pdfBrand}>
                            <div className={styles.brandMark}>ASP</div>
                            <strong>ASP Assist Group</strong>
                        </div>
                        <h2>Daily HAC # {item.number}</h2>
                    </div>

                    <div className={styles.pdfSectionTitle}>WORK DETAIL</div>
                    <table className={styles.pdfTable}>
                        <tbody>
                            <tr><th>Assessment Date</th><td>{formatLongDate(item.date)}</td></tr>
                            <tr><th>Project</th><td>{item.project}</td></tr>
                            <tr><th>Scope of Works</th><td>{item.scopeOfWorks}</td></tr>
                            <tr><th>Work Addresses</th><td>{item.workAddresses}</td></tr>
                            <tr><th>Weather Condition</th><td>{item.weatherCondition}</td></tr>
                            <tr><th>Status</th><td>{item.status}</td></tr>
                        </tbody>
                    </table>

                    <div className={styles.pdfSectionTitle}>EMERGENCY RESPONSE</div>
                    <table className={styles.pdfTable}>
                        <tbody>
                            <tr><th>Emergency Muster Location</th><td>{item.emergencyMusterLocation}</td></tr>
                            <tr><th>First Aid Kit Location</th><td>{item.firstAidKitLocation}</td></tr>
                            <tr><th>Nearest Medical Centre & Contact</th><td>{item.nearestMedicalCentre}</td></tr>
                            <tr><th>Nearest Hospital & Contact</th><td>{item.nearestHospital}</td></tr>
                        </tbody>
                    </table>

                    <div className={styles.pdfSectionTitle}>KEY CONTACTS</div>
                    <table className={styles.pdfTable}>
                        <thead>
                            <tr><th>Role</th><th>Name</th><th>Number</th><th>Company</th></tr>
                        </thead>
                        <tbody>
                            {item.contacts.map(contact => (
                                <tr key={contact.role}>
                                    <td>{contact.role}</td>
                                    <td>{contact.name}</td>
                                    <td>{contact.number}</td>
                                    <td>{contact.company}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>

                    <div className={styles.pdfSectionTitle}>HAZARD ASSESSMENT CHECKLIST</div>
                    <table className={styles.pdfTable}>
                        <thead>
                            <tr><th>Hazard</th><th>Rating</th><th>Control Methods</th><th>Residual Rating</th></tr>
                        </thead>
                        <tbody>
                            {item.hazards.map(hazard => (
                                <tr key={hazard.hazard}>
                                    <td>{hazard.hazard}</td>
                                    <td>{hazard.rating}</td>
                                    <td>{hazard.controlMethods}</td>
                                    <td>{hazard.residualRating}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export default HacPdfPreview;

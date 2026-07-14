import * as React from 'react';
import styles from './TaManagement.module.scss';
import { VERSCHIEBUNG_GRUENDE } from '../models/types';

export interface IModalProps {
    visible: boolean;
    onClose: () => void;
    onConfirm: (neuerTermin: string, grund: string) => void;
}

interface IModalState {
    neuerTermin: string;
    selectedGrund: string;
}

function normalizeIsoDateInput(value: string): string {
    if (!value) return '';
    const match = value.match(/^(\d{0,4})(?:-(\d{0,2}))?(?:-(\d{0,2}))?/);
    if (!match) return '';

    const year = (match[1] || '').slice(0, 4);
    const month = (match[2] || '').slice(0, 2);
    const day = (match[3] || '').slice(0, 2);

    if (!month && !day) return year;
    if (!day) return `${year}-${month}`;
    return `${year}-${month}-${day}`;
}

export default class Modal extends React.Component<IModalProps, IModalState> {
    constructor(props: IModalProps) {
        super(props);
        this.state = {
            neuerTermin: '',
            selectedGrund: ''
        };
    }

    public render(): React.ReactElement<IModalProps> | null {
        if (!this.props.visible) return null;

        return (
            <div className={styles.overlay} onClick={this.props.onClose}>
                <div className={styles.modal} onClick={(e) => e.stopPropagation()}>
                    <h3 className={styles.modalTitle}>Terminverschiebung</h3>

                    <div className={styles.formField}>
                        <label className={styles.formLabel}>Neuer Termin</label>
                        <input
                            type="date"
                            className={styles.dateInput}
                            value={this.state.neuerTermin}
                            onChange={(e) => {
                                const normalized = normalizeIsoDateInput(e.target.value);
                                if (normalized !== e.target.value) {
                                    e.currentTarget.value = normalized;
                                }
                                this.setState({ neuerTermin: normalized });
                            }}
                        />
                    </div>

                    <div className={styles.formField}>
                        <label className={styles.formLabel} style={{ fontWeight: 600 }}>Grund für Verschiebung</label>
                        <div className={styles.radioGroup}>
                            {VERSCHIEBUNG_GRUENDE.map(grund => (
                                <label key={grund} className={styles.radioItem}>
                                    <input
                                        type="radio"
                                        name="verschiebungsgrund"
                                        value={grund}
                                        checked={this.state.selectedGrund === grund}
                                        onChange={() => this.setState({ selectedGrund: grund })}
                                    />
                                    {grund}
                                </label>
                            ))}
                        </div>
                    </div>

                    <div className={styles.modalBtnRow}>
                        <button className={styles.btnSecondary} onClick={this.props.onClose}>
                            Abbrechen
                        </button>
                        <button
                            className={styles.btnPrimary}
                            disabled={!this.state.neuerTermin || !this.state.selectedGrund}
                            onClick={() => {
                                const parts = this.state.neuerTermin.split('-');
                                const formatted = `${parts[2]}.${parts[1]}.${parts[0]}`;
                                this.props.onConfirm(formatted, this.state.selectedGrund);
                                this.setState({ neuerTermin: '', selectedGrund: '' });
                            }}
                        >
                            Übernehmen
                        </button>
                    </div>
                </div>
            </div>
        );
    }
}

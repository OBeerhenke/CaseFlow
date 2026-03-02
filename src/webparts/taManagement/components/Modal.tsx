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
                            onChange={(e) => this.setState({ neuerTermin: e.target.value })}
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
                            ✖ Abbrechen
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
                            ✓ Übernehmen
                        </button>
                    </div>
                </div>
            </div>
        );
    }
}

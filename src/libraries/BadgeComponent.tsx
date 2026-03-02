import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';

export interface IBadgeComponentProps {
    label?: string;
    color?: string;
    badgeClass?: string;
}

const BadgeComponent: React.FC<IBadgeComponentProps> = (props) => {
    const { label, color, badgeClass } = props;

    const style: React.CSSProperties = {};
    if (color) {
        style.backgroundColor = color;
    }

    return (
        <div className={`page-type-badge ${badgeClass || ''}`} style={style}>
            {label}
        </div>
    );
};

export class BadgeWebComponent extends BaseWebComponent {
    public async connectedCallback() {
        const props = this.resolveAttributes();
        ReactDOM.render(<BadgeComponent {...props} />, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}

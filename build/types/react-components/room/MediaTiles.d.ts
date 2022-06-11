export function CreateTile({ label, type, ...rest }: {
    [x: string]: any;
    label: any;
    type: any;
}): JSX.Element;
export namespace CreateTile {
    namespace propTypes {
        const label: PropTypes.Requireable<PropTypes.ReactNodeLike>;
        const type: PropTypes.Requireable<string>;
    }
}
export function MediaTile({ entry, processThumbnailUrl, onClick, onEdit, onShowSimilar, onCopy, onInfo, ...rest }: {
    [x: string]: any;
    entry: any;
    processThumbnailUrl: any;
    onClick: any;
    onEdit: any;
    onShowSimilar: any;
    onCopy: any;
    onInfo: any;
}): JSX.Element;
export namespace MediaTile {
    export namespace propTypes_1 {
        const entry: PropTypes.Validator<object>;
        const processThumbnailUrl: PropTypes.Requireable<(...args: any[]) => any>;
        const onClick: PropTypes.Requireable<(...args: any[]) => any>;
        const onEdit: PropTypes.Requireable<(...args: any[]) => any>;
        const onShowSimilar: PropTypes.Requireable<(...args: any[]) => any>;
        const onCopy: PropTypes.Requireable<(...args: any[]) => any>;
        const onInfo: PropTypes.Requireable<(...args: any[]) => any>;
    }
    export { propTypes_1 as propTypes };
}
import PropTypes from "prop-types";

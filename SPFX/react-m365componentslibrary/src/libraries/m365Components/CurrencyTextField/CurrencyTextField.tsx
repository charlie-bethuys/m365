import * as React from 'react';
import { CSSProperties } from 'react';
import { TextField, IIconProps, IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles } from '@fluentui/react';

export interface ICurrencyTextFieldProps {
    borderless?: boolean;
    className?: string;
    currency?: CurrencyType;
    description?: string;
    disabled?: boolean;
    errorMessage?: string | JSX.Element;
    iconProps?: IIconProps;
    inputClassName?: string;
    label?: string;
    onBlur?: (value: number) => void;
    onGetErrorMessage?: (value: string) => string | JSX.Element | PromiseLike<string | JSX.Element> | undefined;
    readOnly?: boolean;
    style?: CSSProperties;
    styles?: IStyleFunctionOrObject<ITextFieldStyleProps, ITextFieldStyles>;
    value?: number;
}

export enum CurrencyType {
    Euro = "€",
    Dollar = "$",
    Pound = "£"
}


const CurrencyTextField: React.FunctionComponent<ICurrencyTextFieldProps> = (props: ICurrencyTextFieldProps) => {
    const [textFieldValue, setTextFieldValue] = React.useState(props.value.toFixed(2) || (0).toFixed(2));
    const onChangeTextFieldValue = React.useCallback(
        (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
            setTextFieldValue(newValue || '');
        },
        [],
    );
    const onBlurTextField = React.useCallback(
        (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>) => {
            let value: number = parseFloat(event.currentTarget.value.replace(",", "."));
            if (!value) {
                value = 0;
            }
            setTextFieldValue(value.toFixed(2) || (0).toFixed(2));
            value = parseFloat(value.toFixed(2));
            if (props.onBlur) {
                props.onBlur(value);
            }
        },
        [],
    );

    return (
        <TextField
            borderless={props.borderless}
            className={props.className}
            description={props.description}
            disabled={props.disabled}
            errorMessage={props.errorMessage}
            iconProps={props.iconProps}
            inputClassName={props.inputClassName}
            label={props.label}
            onBlur={onBlurTextField}
            onChange={onChangeTextFieldValue}
            onGetErrorMessage={props.onGetErrorMessage}
            readOnly={props.readOnly}
            style={props.style}
            styles={props.styles}
            suffix={props.currency}
            validateOnFocusOut
            value={textFieldValue}
        />
    );
};

export { CurrencyTextField };
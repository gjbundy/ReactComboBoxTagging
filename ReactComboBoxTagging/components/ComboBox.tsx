import * as React from "react";
import { Option, Theme, Combobox, makeStyles, shorthands, tokens, useId, Tag, useComboboxFilter, FluentProvider } from "@fluentui/react-components"
import type { ComboboxProps, ForwardRefComponent, OptionProps } from "@fluentui/react-components";

const useState = React.useState;

const useStyles = makeStyles({
    root: {
        // Stack the label above the field with a gap
        display: "grid",
        gridTemplateRows: "repeat(1fr)",
        justifyItems: "start",
        ...shorthands.gap("2px"),
        maxWidth: "400px",
    },
    tagsList: {
        listStyleType: "none",
        marginBottom: tokens.spacingVerticalXXS,
        marginTop: 0,
        paddingLeft: 0,
        display: "flex",
        gridGap: tokens.spacingHorizontalXS,
    },
    tagContainer: {
        display: "flex",
        flexDirection: "row",
        justifyContent: "flex-start",
        flexWrap: "wrap",
        width: "100%", // Full width
        minHeight: "32px", // Minimum height to start with
        flexGrow: 1, // Used to make the container grow based on content
    },
});

export interface IComboBoxTagPickerProps {
    thisSelectedOption: string | undefined;
    availableOptions: typeof Option[];
    theme: Theme;
    initialSelectedOptionsString?: string;
    tagOptionsFromTable: string[];
    //onOptionSelect?: (event: any, data: any) => void;
    onSelectedOptionsChanged?: (selectedOptions: string[]) => void;
}

export const ComboboxTagPicker = React.memo((props: IComboBoxTagPickerProps) => {
    const { thisSelectedOption, availableOptions, theme, tagOptionsFromTable } = props;
    const comboId = useId("combo-multi");
    const selectedListId = `${comboId}-selection`;
    const [inputValue, setInputValue] = React.useState('');

    // refs for managing focus when removing tags
    const selectedListRef = React.useRef<HTMLUListElement>(null);
    const comboboxInputRef = React.useRef<HTMLInputElement>(null);

    //define temporary options for testing
    //const options = ["Garrett", "Aadi", "DJ", "Jamie", "Stephanie", "Lindsey", "Louise", "Tanner",]

    const options = tagOptionsFromTable;
    //console.log("options: ", options)

    const [selectedOptions, setSelectedOptions] = useState<string[]>([]);
    const styles = useStyles();

    // set initial selected options
    React.useEffect(() => {
        if (props.initialSelectedOptionsString) {
            //split the string into an array of options
            const initialOptions = props.initialSelectedOptionsString.split(",").map((o) => o.trim());
            setSelectedOptions(initialOptions);
            if (props.onSelectedOptionsChanged) {
                props.onSelectedOptionsChanged(initialOptions);
            }
        }
    }, []); // Empty dependency array means this effect runs once on mount


    const onSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
        console.log("onSelect: ", data.selectedOptions);
        setSelectedOptions(data.selectedOptions);
        if (props.onSelectedOptionsChanged) {
            props.onSelectedOptionsChanged(data.selectedOptions);
        }
    };

    const onTagClick = (option: string, index: number) => {
        // remove selected option
        const updatedOptions = selectedOptions.filter((o) => o !== option);
        setSelectedOptions(updatedOptions);
        if (props.onSelectedOptionsChanged) {
            props.onSelectedOptionsChanged(updatedOptions);
        }
    };

    const labelledBy =
        selectedOptions.length > 0 ? `${comboId} ${selectedListId}` : comboId;

    return (
        <FluentProvider theme={props.theme}>
            <div className={styles.root}>
                {/* <label id={comboId}>Tags for this record</label> */}
                <div className={styles.tagContainer}>
                    <FluentProvider theme={props.theme}>
                        {selectedOptions.length ? (
                            selectedOptions.map((option, i) => (
                                <Tag key={i}
                                    appearance="brand"
                                    shape="circular"
                                    dismissible
                                    dismissIcon={{ "aria-label": "remove" }}
                                    onClick={() => onTagClick(option, i)}>
                                    {option}
                                </Tag>
                            ))) : null}
                    </FluentProvider>
                </div>
                <Combobox
                    appearance="outline"
                    aria-labelledby={labelledBy}
                    multiselect={true}
                    placeholder="Select one or more tags"
                    selectedOptions={selectedOptions}
                    onOptionSelect={onSelect}
                    ref={comboboxInputRef}
                    {...props}
                >
                    {options.map((option) => (
                        <Option key={option}>{option}</Option>
                    ))}
                </Combobox>
            </div>
        </FluentProvider>

    );
});
ComboboxTagPicker.displayName = "ComboboxTagPicker";


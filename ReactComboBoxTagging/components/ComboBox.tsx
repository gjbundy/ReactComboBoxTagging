import * as React from "react";
import { Option, Theme, Combobox, makeStyles, shorthands, tokens, useId, Tag, useComboboxFilter, FluentProvider, Button } from "@fluentui/react-components"
import type { ComboboxProps } from "@fluentui/react-components";

const useState = React.useState;

const useStyles = makeStyles({
    root: {
        // Stack the label above the field with a gap
        display: "grid",
        gridTemplateRows: "repeat(1fr)",
        justifyItems: "start",
        ...shorthands.gap("2px"),
        //maxWidth: "400px",
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
    comboBoxAndButton:{
        display: "flex",
        flexDirection: "row",
        justifyContent: "flex-start",
        flexWrap: "wrap",
        width: "100%", // Full width
        minHeight: "32px", // Minimum height to start with
        flexGrow: 1, // Used to make the container grow based on content
        ...shorthands.gap("10px"),
    }
});

export interface IComboBoxTagPickerProps extends ComboboxProps {
    thisSelectedOption: string | undefined;
    availableOptions: typeof Option[];
    theme: Theme;
    initialSelectedOptionsString?: string;
    tagOptionsFromTable: string[];
    onSelectedOptionsChanged?: (selectedOptions: string[]) => void;
}

export const ComboboxTagPicker = React.memo((props: IComboBoxTagPickerProps) => {
    const { thisSelectedOption, availableOptions, theme, tagOptionsFromTable } = props;
    const comboId = useId("combo-multi");
    const selectedListId = `${comboId}-selection`;
    const [inputValue, setInputValue] = React.useState(''); //added for filtering implementation
    const [newTags, setNewTags] = React.useState<string[]>([]); //added for tracking new tags
    const newTagStyle = { backgroundColor: 'green', color: 'white' }; //added for new tag styling

    // refs for managing focus when removing tags
    const selectedListRef = React.useRef<HTMLUListElement>(null);
    const comboboxInputRef = React.useRef<HTMLInputElement>(null);

    // Handler to reset search/filter when Combobox loses focus
    const handleBlur = () => {
        setInputValue(''); // Reset the inputValue state
    };

    //Testing to see if I can filter the options
    const filteredOptions = useComboboxFilter(inputValue, tagOptionsFromTable, {
        noOptionsMessage: 'No tags match your search. Press enter to add a new tag.',
    });

    const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
        if (event.key === 'Enter' && inputValue.trim() !== '' && !selectedOptions.includes(inputValue)) {
            // Add the new tag (current inputValue) to selectedOptions
            const newSelectedOptions = [...selectedOptions, inputValue.trim()];
            setSelectedOptions(newSelectedOptions);

            // Add the new tag to the newTags state
            setNewTags((prevNewTags) => [...prevNewTags, inputValue.trim()]);


            // Optionally, call onSelectedOptionsChanged if it needs to trigger external actions
            if (props.onSelectedOptionsChanged) {
                props.onSelectedOptionsChanged(newSelectedOptions);
            }

            // Reset inputValue to clear the combobox input field
            setInputValue('');
        }
    };


    //removed for filtering implementation:
    //const options = tagOptionsFromTable;
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
        if(newTags.includes(option)){
            const updatedNewTags = newTags.filter((o) => o !== option);
            setNewTags(updatedNewTags);
        }
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
                                    onClick={() => onTagClick(option, i)}
                                    style={newTags.includes(option) ? newTagStyle : {}}
                                    >
                                    {option}
                                </Tag>
                            ))) : null}
                    </FluentProvider>
                </div>
                <div className={styles.comboBoxAndButton}>
                    <Combobox
                        appearance="outline"
                        aria-labelledby={labelledBy}
                        multiselect={true}
                        placeholder="Select one or more tags"
                        selectedOptions={selectedOptions}
                        onOptionSelect={onSelect}
                        onChange={(ev) => setInputValue(ev.target.value)} //added for filtering implementation
                        onBlur={handleBlur} //Added to reset search/filter when Combobox loses focus
                        onKeyDown={handleKeyDown} //Added for save on enter
                        ref={comboboxInputRef}
                        {...props}
                    >
                        {/* {options.map((option) => (
                            <Option key={option}>{option}</Option>
                        ))} */}
                        {filteredOptions}
                    </Combobox>
                    <Button
                        appearance="primary"
                        disabled={newTags.length === 0}
                        onClick={() => {
                            console.log("selectedOptions: ", selectedOptions);
                            console.log("newTags: ", newTags);
                        }}
                    >
                        Save New Tags
                    </Button>
                </div>
            </div>
        </FluentProvider>

    );
});
ComboboxTagPicker.displayName = "ComboboxTagPicker";


import * as React from "react";
import { useForceUpdate } from '@fluentui/react-hooks';
import { IInputs } from "../generated/ManifestTypes";
import { Option, Combobox, Theme, Spinner, makeStyles, shorthands, tokens, useId, Tag, useComboboxFilter, FluentProvider, Button } from "@fluentui/react-components"
import type { ComboboxProps } from "@fluentui/react-components";

const useState = React.useState;


const useStyles = makeStyles({
    root: {
        // Stack the label above the field with a gap
        display: "grid",
        gridTemplateRows: "repeat(1fr)",
        justifyItems: "start",
        ...shorthands.gap("2px"),
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
    comboBoxAndButton: {
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
    context: ComponentFramework.Context<IInputs>;
    tableName: string;
    thisSelectedOption: string | undefined;
    availableOptions: typeof Option[];
    theme: Theme;
    initialSelectedOptionsString?: string;
    onSelectedOptionsChange: (selectedOptionsString: string) => void;
    onSaveTags?: (newTags: string[]) => void;
}

export const ComboboxTagPicker = React.memo((props: IComboBoxTagPickerProps) => {
    const { thisSelectedOption, availableOptions, theme, onSaveTags, context, tableName, onSelectedOptionsChange } = props;
    const [optionsFromTable, setOptionsFromTable] = useState<string[]>([]);
    const styles = useStyles();
    const forceUpdate = useForceUpdate();
    const comboId = useId("combo-multi");
    const selectedListId = `${comboId}-selection`;
    const [inputValue, setInputValue] = React.useState(''); //added for filtering implementation
    const [newTags, setNewTags] = React.useState<string[]>([]); //added for tracking new tags
    const [isComponentLoading, setIsLoading] = React.useState<boolean>(true);
    const [selectedOptions, setSelectedOptions] = useState<string[]>([]);
    const newTagStyle = { backgroundColor: 'green', color: 'white' }; //added for new tag styling

    //get the existing tags
    const initialTags = props.context.parameters.Tags.raw || undefined;
    console.log("Initial tags in ComboBox are: " + initialTags);

    // refs for managing focus when removing tags
    const selectedListRef = React.useRef<HTMLUListElement>(null);
    const comboboxInputRef = React.useRef<HTMLInputElement>(null);

    // Handler to reset search/filter when Combobox loses focus
    const handleBlur = () => {
        setInputValue(''); // Reset the inputValue state
    };

    const filteredOptions = useComboboxFilter(inputValue, optionsFromTable, {
        noOptionsMessage: 'No tags match your search. Press enter to add a new tag.',
    });

    // Convert selectedOptions array to string
    const selectedOptionsString = selectedOptions.join(',');

    const handleKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
        if (event.key === 'Enter' && inputValue.trim() !== '' && !selectedOptions.includes(inputValue)) {
            // Add the new tag (current inputValue) to selectedOptions
            const newSelectedOptions = [...selectedOptions, inputValue.trim()];
            setSelectedOptions(newSelectedOptions);

            // Add the new tag to the newTags state
            setNewTags((prevNewTags) => [...prevNewTags, inputValue.trim()]);
            // Reset inputValue to clear the combobox input field
            setInputValue('');
        }
    };

    React.useEffect(() => {
        const getOptionsFromTable = async () => {
            console.log("inside the useEffect for getOptionsFromTable. Starting async call to retrieveMultipleRecords, with tableName: ", tableName);
            const result = await props.context.webAPI.retrieveMultipleRecords(tableName);
            console.log("result: ", result.entities.map(entity => entity['dtapps_tagtext']));
            setOptionsFromTable(result.entities.map(entity => entity['dtapps_tagtext']));
        };
        getOptionsFromTable().then(() => {
            console.log("inside of the then clause for getOptionsFromTable, with initialTags: ", initialTags);
            if(initialTags) {
                const initialOptions = initialTags.split(",").map((o) => o.trim());
                setSelectedOptions(initialOptions);
            }
            setIsLoading(false);
        });
    }, [context, tableName, initialTags]);

    React.useEffect(() => {
        onSelectedOptionsChange(selectedOptionsString);
    }, [selectedOptionsString]);

    const onSelect: ComboboxProps["onOptionSelect"] = (event, data) => {
        setSelectedOptions(data.selectedOptions);
    };

    const onTagClick = (option: string, index: number) => {
        // remove selected option
        const updatedOptions = selectedOptions.filter((o) => o !== option);
        if (newTags.includes(option)) {
            const updatedNewTags = newTags.filter((o) => o !== option);
            setNewTags(updatedNewTags);
        }
        //Need to see if i have to update the string here as well
    };


    const labelledBy =
        selectedOptions.length > 0 ? `${comboId} ${selectedListId}` : comboId;

    if (isComponentLoading) {
        return (
            React.createElement(
                Spinner, { label: "Fetching Tags..." }
            ));
    } else {
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
                            {filteredOptions}
                        </Combobox>
                        <Button
                            appearance="primary"
                            disabled={newTags.length === 0}
                            onClick={() => {
                                console.log("selectedOptions: ", selectedOptions);
                                console.log("newTags: ", newTags);
                                // { handleSaveClicked }
                            }}
                        >
                            Save New Tags
                        </Button>
                    </div>
                </div>
            </FluentProvider>

        );
    }
});
ComboboxTagPicker.displayName = "ComboboxTagPicker";


import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ComboboxTagPicker, IComboBoxTagPickerProps } from "./components/ComboBox";
import { lightThemeCompanyBlue, darkThemeCompanyBlue } from "./customthemes/CompanyBlue";
import * as React from "react";
import { teamsLightTheme, teamsDarkTheme, webLightTheme, webDarkTheme, Option, Theme, Spinner, } from "@fluentui/react-components";

export class ReactComboBoxTagging implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private state: ComponentFramework.Dictionary;
    private availableOptions: typeof Option[];
    private thisSelectedOption: string | undefined;
    private selectedOptionsOutput: string | undefined;
    private themeSelected: Theme;
    private tagOptionsFromTable: string[];
    private loadedDataDone: boolean = false;
    private tableName: string;
    private initialSelectedTags: string[];

    /**
     * Empty constructor.
     */
    constructor() {
        // Bind the handleSelectedOptions method to 'this'
        // this.handleSelectedOptions = this.handleSelectedOptions.bind(this);
        this.saveTags = this.saveTags.bind(this);
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        _context: ComponentFramework.Context<IInputs>,
        _notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
    ){
        this.notifyOutputChanged = _notifyOutputChanged;
        this.context = _context;
        this.state = _state || {};
        var _themeSelected = this.context.parameters.Theme.raw;

        var _selectedOptionsOutput = this.context.parameters.Tags.raw;
        this.tableName = this.context.parameters.TagsDB.raw || '';
        this.selectedOptionsOutput = _selectedOptionsOutput ?? undefined;

        if (this.tableName === undefined) {
            this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            this.loadedDataDone = true;
        } else {
            //Set temporary Values when testing locally only
            // this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            // this.loadedDataDone = true;
            // this.notifyOutputChanged();
        }

        switch (_themeSelected) {
            case "Web Light Theme":
                this.themeSelected = webLightTheme
                break
            case "Web Dark Theme":
                this.themeSelected = webDarkTheme
                break
            case "Teams Light Theme":
                this.themeSelected = teamsLightTheme
                break
            case "Teams Dark Theme":
                this.themeSelected = teamsDarkTheme
                break
            case "Use Platform Theme":
                this.themeSelected = this.context.fluentDesignLanguage?.tokenTheme
                break
            case "Company Blue Light Theme":
                this.themeSelected = lightThemeCompanyBlue
                break
            case "Company Blue Dark Theme":
                this.themeSelected = darkThemeCompanyBlue
                break
            default:
                this.themeSelected = lightThemeCompanyBlue
                break
        }
        this.selectedOptionsOutput = _selectedOptionsOutput ?? undefined;
    }

    private handleSelectedOptions(options: string[]): void {
        console.log("In the handleSelectedOptions method: ", options);
        this.selectedOptionsOutput = options.join(",");
        console.log("handleSelectedOptions - The selectedOptionsOutput in handleSelectedOptions is: ", this.selectedOptionsOutput)
        // this.notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {

        // Retrieve the value of themeSelected from above
        let themeSelected: Theme = this.themeSelected as unknown as Theme;

        return React.createElement(ComboboxTagPicker, {
            availableOptions: this.availableOptions,
            thisSelectedOption: this.thisSelectedOption,
            initialSelectedOptionsString: this.selectedOptionsOutput,
            theme: themeSelected,
            context: this.context,
            tableName: this.tableName,
            onSelectedOptionsChanged: this.handleSelectedOptions,
            onSaveTags: this.saveTags
        } as IComboBoxTagPickerProps);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        // console.log("The selectedOptionsOutput in the getOutputs method is: ", this.selectedOptionsOutput)
        return { Tags: this.selectedOptionsOutput } as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    private saveTags(newTagsArray: string[]): void {
        const tableName = this.context.parameters.TagsDB.raw; // Assuming 'TagsDB' is the parameter name for the table
        const newTags = newTagsArray; // Assuming this.newTags holds the tags to save

        if (!tableName || !newTags || newTags.length === 0) {
            console.warn("No tags to save");
            return;
        } else {
            newTags.forEach(tag => {
                console.log("Saving tag: ", tag);
                const data = {
                    "dtapps_tagtext": tag
                };

                this.context.webAPI.createRecord(tableName, data)
                    .then(result => {
                        console.log(`Tag saved successfully: ${result.id}`);
                        // Optionally, perform additional actions upon success, e.g., clear the newTags array or refresh the view
                        // this.getOptionsFromTable(tableName).then((options) => {
                        //     console.log("Data received from the table after saving: ", options);
                        //     this.tagOptionsFromTable = options;
                        //     this.loadedDataDone = true;
                        //     this.notifyOutputChanged();
                        //     this.updateView(this.context);
                        // }).catch((error) => {
                        //     console.error("Error fetching data from the table after saving: ", error);

                        // });
                    })
                    .catch(error => {
                        console.error("Error saving the tag", error);
                        // Handle error, e.g., show an error message to the user
                    });
            });
        }
    }
}

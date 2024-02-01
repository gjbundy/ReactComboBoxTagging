import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ComboboxTagPicker, IComboBoxTagPickerProps } from "./components/ComboBox";
import { lightThemeCompanyBlue, darkThemeCompanyBlue } from "./customthemes/CompanyBlue";
import * as React from "react";
import { teamsLightTheme, teamsDarkTheme, webLightTheme, webDarkTheme, Option, Skeleton, Theme, } from "@fluentui/react-components";
import type { SkeletonProps } from "@fluentui/react-components";

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



    /**
     * Empty constructor.
     */
    constructor() {
        // Bind the handleSelectedOptions method to 'this'
        this.handleSelectedOptions = this.handleSelectedOptions.bind(this)
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public async init(
        _context: ComponentFramework.Context<IInputs>,
        _notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary
    ): Promise<void> {
        this.notifyOutputChanged = _notifyOutputChanged;
        this.context = _context;
        this.state = _state || {};
        var _themeSelected = this.context.parameters.Theme.raw;
        var _selectedOptionsOutput = this.context.parameters.Tags.raw;
        var _tagOptionsFromTable: string[] = [];

        console.log("In the init method");

        // Get the table name from the context parameters
        const tableName = this.context.parameters.TagsDB.raw || '';
        console.log("Table Name: ", tableName);

        if (tableName === undefined) {
            this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            this.loadedDataDone = true;
            this.notifyOutputChanged();
        } else {
            // Fetch the data from the table asynchronously
            // console.log("Fetching data from the table: ", tableName)
            // await this.getOptionsFromTable(tableName).then((options) => {
            //     console.log("Data received from the table first: ", options);
            //     _tagOptionsFromTable = options;
            //     this.tagOptionsFromTable = _tagOptionsFromTable;
            //     this.loadedDataDone = true;
            //     this.notifyOutputChanged();            
            // }).catch((error) => {
            //     console.error("Error fetching data from the table: ", error);
            //     // Handle error appropriately
            // });

            //Set temporary Values
            this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            this.loadedDataDone = true;
            this.notifyOutputChanged();
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
        console.log("handleSelectedOptions - The selectedOptionsOutput is: ", this.selectedOptionsOutput)
        this.notifyOutputChanged();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {

        console.log("Starting the updateView method and the loadedDataDone is: ", this.loadedDataDone);
        if (!this.loadedDataDone) {
            return React.createElement(
                Skeleton, { width: "100%" }
            )
        }

        // Retrieve the value of themeSelected from above
        let themeSelected: Theme = this.themeSelected as unknown as Theme;

        //retrieve the value of tagOptionsFromTable from above
        let tagOptionsFromTablePassed: string[] = this.tagOptionsFromTable;

        console.log("In the updateVew method: ", tagOptionsFromTablePassed);
        return React.createElement(ComboboxTagPicker, {
            availableOptions: this.availableOptions,
            thisSelectedOption: this.thisSelectedOption,
            initialSelectedOptionsString: this.selectedOptionsOutput,
            theme: themeSelected,
            tagOptionsFromTable: tagOptionsFromTablePassed,
            //onOptionSelect: this.selectedOptionsOutput,
            onSelectedOptionsChanged: this.handleSelectedOptions
        } as IComboBoxTagPickerProps);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return { Tags: this.selectedOptionsOutput } as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    private async getOptionsFromTable(tableName: string): Promise<string[]> {
        const result = await this.context.webAPI.retrieveMultipleRecords(tableName);
        console.log("Data being returned from the table: ", result.entities.map(entity => entity['dtapps_tagtext']));
        return result.entities.map(entity => entity['dtapps_tagtext']);

    }
}

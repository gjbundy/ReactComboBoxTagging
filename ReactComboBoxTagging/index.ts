import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ComboboxTagPicker, IComboBoxTagPickerProps } from "./components/ComboBox";
import { lightThemeCompanyBlue, darkThemeCompanyBlue } from "./customthemes/CompanyBlue";
import * as React from "react";
import { teamsLightTheme, teamsDarkTheme, webLightTheme, webDarkTheme, Option, Theme } from "@fluentui/react-components";

export class ReactComboBoxTagging implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private multiSelect: boolean;
    private state: ComponentFramework.Dictionary;
    private availableOptions: typeof Option[];
    private selectedOptions: string | undefined;
    private selectedOptionsOutput: string | undefined;
    private themeSelected: Theme;
    private tagShape: string;
    private tagAppearance: string;
    private tagOptionsFromTable: string[];
    private loadedDataDone: boolean = false;
    private tableName: string;

    /**
     * Empty constructor.
     */
    constructor() {
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
    ) {
        this.notifyOutputChanged = _notifyOutputChanged;
        this.context = _context;
        this.state = _state || {};
        var _themeSelected = this.context.parameters.Theme.raw;
        var _tagShape = this.context.parameters.TagShape.raw;
        var _tagAppearance = this.context.parameters.TagAppearance.raw;
        var _controlType = this.context.parameters.ControlType.raw;
        this.tableName = this.context.parameters.TagsDB.raw || '';

        if (this.tableName === undefined) {
            this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            this.loadedDataDone = true;
        } else {
            //Set temporary Values when testing locally only
            // this.tagOptionsFromTable = ["No Options Retrieved", "Test 1", "Test 2"];
            // this.loadedDataDone = true;
            // this.notifyOutputChanged();
        }

        switch (_controlType) {
            case "Multiple Selections":
                this.multiSelect = true;
                break;
            case "Single Selection Only": //User has selected single-Select Optiopn
                this.multiSelect = false;
                break;
            default: 
                this.multiSelect = false;
                break;
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
        switch (_tagShape) {
            case "Circular":
                this.tagShape = "circular"
                break
            case "Rounded":
                this.tagShape = "rounded"
                break
            default:
                this.tagShape = "circular"
                break
        }
        switch (_tagAppearance) {
            case "Brand":
                this.tagAppearance = "brand"
                break
            case "Filled":
                this.tagAppearance = "filled"
                break
            case "Outline":
                this.tagAppearance = "outline"
                break
            default:
                this.tagAppearance = "brand"
                break
        }
    }

    private handleSelectedOptionsChange = (selectedOptionsString: string) => {
        this.selectedOptions = selectedOptionsString;
        this.notifyOutputChanged();
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
            context: this.context,
            disabled: context.mode.isControlDisabled,
            initialSelectedOptionsString: this.selectedOptionsOutput,
            multiSelect: this.multiSelect,
            onSelectedOptionsChange: this.handleSelectedOptionsChange.bind(this),
            onSaveTags: this.saveTags.bind(this),
            tableName: this.tableName,
            tagAppearance: this.tagAppearance,
            tagShape: this.tagShape,
            theme: themeSelected,
            thisSelectedOption: this.selectedOptions,
        } as IComboBoxTagPickerProps);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return { Tags: this.selectedOptions } as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }

    private saveTags(newTagsArray: string[]): void {
        const tableName = this.context.parameters.TagsDB.raw;
        const newTags = newTagsArray;

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
                        this.updateView(this.context);
                    })
                    .catch(error => {
                        console.error("Error saving the tag", error);
                    });
            });
        }
    }
}

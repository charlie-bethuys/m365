import * as React from 'react';
import { DefaultButton, PrimaryButton, StackItem, Checkbox, Panel, PanelType, SearchBox, Stack, IconButton, CommandBarButton } from '@fluentui/react';
import { M365ComponentsLibrary, PropertyType } from '../M365ComponentsLibrary';
import styles from './FilterPanel.module.scss';

export interface IFilterPanelProps {
    isOpen: boolean;
    items: any[];
    property: string;
    type?: PropertyType;
    title?: string;
    defaultSelectedValues?: any[];
    onPropertyFilterChange?: (selectedValues: string[]) => void;
    onDismissed?: () => void;
}

const PropertyFilterPanel: React.FunctionComponent<IFilterPanelProps> = (props: IFilterPanelProps) => {
    const [isOpen, setIsOpen] = React.useState(props.isOpen);
    const [searchBoxText, setSearchBoxText] = React.useState("");
    const [selectedFilterValues, setSelectedFilterValues] = React.useState(props.defaultSelectedValues ? props.defaultSelectedValues : []);


    const getFilterValues = (items: any[], property: string): any[] => {
        let propertyValues: any[] = [];
        items.forEach(
            item => {
                const value = M365ComponentsLibrary.getPropertyValueByPath(item, property);
                if (value && propertyValues.indexOf(value) === -1) {
                    propertyValues.push(value as string);
                }
            }
        );

        return propertyValues.sort();
    };

    const onRemoveSelectedFilterValues = (): void => {
        setSelectedFilterValues([]);
    };

    const onApplySelectedFilterValues = (): void => {
        props.onPropertyFilterChange(selectedFilterValues);
    };

    const filterValues = getFilterValues(props.items, props.property);

    return (
        <Panel
            headerText={props.title ? props.title : `Filtrer par « ${props.property} »`}
            isBlocking={false}
            type={PanelType.smallFixedFar}
            isOpen={isOpen}
            onDismiss={props.onDismissed}
            className={styles.filterPanel}>
            <Stack verticalAlign={"space-evenly"}>
                <StackItem>
                    <SearchBox
                        className={styles.searchBox}
                        placeholder="Rechercher des filtres..."
                        onChange={
                            (event: React.ChangeEvent<HTMLInputElement>, newValue: string) => {
                                setSearchBoxText(newValue);
                            }
                        } />
                </StackItem>
                <StackItem>
                    <Stack verticalAlign={"space-evenly"}>
                        {
                            filterValues.map(
                                (filterValue, index) => {
                                    let filterValueText = "";
                                    switch (props.type) {
                                        case "date":
                                            filterValueText = M365ComponentsLibrary.formatDate(filterValue);
                                            break;
                                        case "datetime":
                                            filterValueText = M365ComponentsLibrary.formatDateHeure(filterValue);
                                            break;
                                        case "boolean":
                                            filterValueText = filterValue ? "Vrai" : "Faux";
                                            break;
                                        default:
                                            filterValueText = filterValue.toString();
                                    }
                                    if (M365ComponentsLibrary.match(filterValueText, searchBoxText)) {
                                        return <Checkbox
                                            key={"FilterValue" + index.toString()}
                                            className={styles.filterValue}
                                            label={filterValueText}
                                            checked={selectedFilterValues.indexOf(filterValue) != -1}
                                            onChange={
                                                (ev, checked) => {
                                                    let selectedValues = selectedFilterValues;
                                                    if (checked) {
                                                        if (selectedValues.indexOf(filterValue) === -1) {
                                                            selectedValues.push(filterValue);
                                                        }
                                                    } else {
                                                        let i = selectedValues.indexOf(filterValue);
                                                        if (i != -1) {
                                                            selectedValues.splice(i, 1);
                                                        }
                                                    }
                                                    setSelectedFilterValues([...selectedValues]);
                                                }
                                            } />;
                                    }
                                }
                            )
                        }
                    </Stack>
                </StackItem>
                <StackItem>
                    <Stack horizontal horizontalAlign={"center"} className={styles.buttonsRow}>
                        <PrimaryButton text={"Appliquer"} onClick={onApplySelectedFilterValues} />
                        <DefaultButton text={"Effacer tout"} onClick={onRemoveSelectedFilterValues} />
                    </Stack>
                    <Stack horizontal horizontalAlign={"center"}>
                        <DefaultButton text={"Fermer"} onClick={props.onDismissed} />
                    </Stack>
                </StackItem>
            </Stack>
        </Panel>
    );
};


const FilterPanel: React.FunctionComponent<IFilterPanelProps> = (props: IFilterPanelProps) => {
    return (
        <PropertyFilterPanel {...props} />
    );
};

export default FilterPanel;
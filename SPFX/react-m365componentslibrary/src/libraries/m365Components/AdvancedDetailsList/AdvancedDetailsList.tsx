import * as React from 'react';
import {
    ColumnActionsMode, CommandBar, ContextualMenu, DetailsList, DetailsListLayoutMode, DirectionalHint, IColumn,
    ICommandBarItemProps, IContextualMenuProps, IGroup, IPanelProps, SelectionMode, Panel, PanelType,
    ResizeGroupDirection, Selection
} from '@fluentui/react';
import FilterPanel, { IFilterPanelProps } from '../FilterPanel/FilterPanel';
import { M365ComponentsLibrary, ISortBy, PropertyType } from '../M365ComponentsLibrary';
import * as xlsx from "xlsx";
import { saveAs } from "file-saver";

export interface IAdvancedDetailsListProps {
    items: any[];
    itemKeyProperty?: string;
    defaultView: IView;
    disableFilterBy?: boolean;
    disableSortBy?: boolean;
    disableGroupBy?: boolean;
    disableCommandBar?: boolean;
    disableCommandBarNew?: boolean;
    disableCommandBarEdit?: boolean;
    disableCommandBarRemove?: boolean;
    disableCommandBarExport?: boolean;
    formPanelType?: PanelType;
    selectionMode?: SelectionMode;
    showFormPanel?: boolean;
    onNew?: () => any;
    onEdit?: (item: any) => any;
    onRemove?: () => any;
    onExport?: () => any;
    onSelectionChanged?: (items: any[], selection: Selection) => any;
    onItemInvoked?: (items: any, index: number) => any;
    onRemoveItems?: (items: any[]) => any;
    onRenderFormPanel?: (item: any, closePanel: () => void) => JSX.Element;
    className?: string;
    compact?: boolean;
}

export interface IAdvancedDetailsListState {
    items: any[];
    groups: IGroup[];
    columns: IColumn[];
    view: IView;
}

export interface IField {
    displayName: string;
    fieldName: string;
    type?: PropertyType;
    isSortable?: boolean;
    isGroupable?: boolean;
    isFilterable?: boolean;
    isHidden?: boolean;
    isIconOnly?: boolean;
    isResizable?: boolean;
    minWidth?: number;
    maxWidth?: number;
    iconName?: string;
    onRender?: (item?: any, index?: number, column?: IColumn) => any;
}

export interface IFilterBy {
    property: string;
    values: string[];
}

export interface IGroupBy {
    fieldName: string;
    isCollapsed: boolean;
}

export interface IView {
    fields: IField[];
    filteredBy?: IFilterBy[];
    sortedBy?: ISortBy[];
    groupedBy?: IGroupBy[];
}

export interface IFormPanelProps {
    panelProps: IPanelProps;
    selectedItem: any;
}

const AdvancedDetailsList: React.FunctionComponent<IAdvancedDetailsListProps> = (props: IAdvancedDetailsListProps) => {
    const selection = new Selection({
        getKey: getKey,
        onSelectionChanged: () => {
            if (props.onSelectionChanged) {
                props.onSelectionChanged([...selection.getSelection()], selection);
            }
            if (!props.disableCommandBar) {
                setCommandBarItemProps(getCommandBarItemProps);
            }
        }
    });
    const [state, dispatch] = React.useReducer(reducer, props, init);
    const [commandBarItemProps, setCommandBarItemProps] = React.useState(getCommandBarItemProps);
    const [contextualMenuProps, setContextualMenuProps] = React.useState(undefined as IContextualMenuProps);
    const [filterPanelProps, setFilterPanelProps] = React.useState(undefined as IFilterPanelProps);
    const [formPanelProps, setFormPanelProps] = React.useState(undefined as IFormPanelProps);
    React.useEffect(() => { dispatch(state.view); }, [props.items]);

    function filter(items: any[], filterProperties: IFilterBy[]) {
        let filteredItems: any[] = [];
        if (filterProperties && filterProperties.length > 0) {
            items.forEach(
                item => {
                    let match = true;
                    for (let i = 0; i < filterProperties.length; i++) {
                        match = match && filterProperties[i].values.indexOf(M365ComponentsLibrary.getPropertyValueByPath(item, filterProperties[i].property)) != -1;
                    }
                    if (match) {
                        filteredItems.push(item);
                    }
                }
            );
        } else {
            filteredItems = items;
        }
        return filteredItems;
    }

    function getColumns(view: IView): IColumn[] {
        let columns: IColumn[] = [];
        view.fields.forEach((field) => {
            if (!field.isHidden) {
                const isFiltered = view.filteredBy.filter(filteredByItem => filteredByItem.property === field.fieldName).length > 0;
                const sortBy = view.sortedBy.filter(sortedByItem => sortedByItem.property === field.fieldName);
                const isGrouped = view.groupedBy.filter(groupedByItem => groupedByItem.fieldName === field.fieldName).length > 0;
                columns.push({
                    key: field.fieldName,
                    name: field.displayName,
                    fieldName: field.fieldName,
                    isResizable: field.isResizable,
                    minWidth: field.minWidth ? field.minWidth : 100,
                    maxWidth: field.maxWidth ? field.maxWidth : 350,
                    onRender: (item) => {
                        if (field.onRender) {
                            return field.onRender(item);
                        }
                        else {
                            return M365ComponentsLibrary.formatValueAsString(field.type, M365ComponentsLibrary.getPropertyValueByPath(item, field.fieldName));
                        }
                    },
                    onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
                        if (column.columnActionsMode !== ColumnActionsMode.disabled) {
                            setContextualMenuProps(getContextualMenuProps(ev, field));
                        }
                    },
                    onColumnContextMenu: (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
                        if (column.columnActionsMode !== ColumnActionsMode.disabled) {
                            setContextualMenuProps(getContextualMenuProps(ev, field));
                        }
                    },
                    isIconOnly: field.isIconOnly,
                    iconName: field.iconName,
                    isFiltered: isFiltered,
                    isGrouped: isGrouped,
                    isSorted: sortBy && sortBy.length > 0 ? true : false,
                    isSortedDescending: sortBy && sortBy.length > 0 ? sortBy[0].isSortedDescending : undefined,
                    columnActionsMode: ColumnActionsMode.hasDropdown,
                    data: field.type
                });
            }
        });
        return columns;
    }

    function getCommandBarItemProps(): ICommandBarItemProps[] {
        let tmpCommandBarItemProps: ICommandBarItemProps[] = [];
        if (!props.disableCommandBarNew) {
            tmpCommandBarItemProps.push({
                key: "new",
                name: "Nouveau",
                iconProps: {
                    iconName: "Add"
                },
                disabled: false,
                onClick: onNew
            });
        }
        if (!props.disableCommandBarEdit) {
            tmpCommandBarItemProps.push({
                key: "edit",
                name: "Modifier",
                iconProps: {
                    iconName: "Edit"
                },
                disabled: selection.getSelection().length !== 1,
                onClick: onEdit
            });
        }
        if (!props.disableCommandBarRemove) {
            tmpCommandBarItemProps.push({
                key: "remove",
                name: "Supprimer",
                iconProps: {
                    iconName: "Cancel"
                },
                disabled: selection.getSelection().length === 0,
                onClick: onRemove
            });
        }
        if (!props.disableCommandBarExport) {
            tmpCommandBarItemProps.push({
                key: "export",
                name: "Exporter vers Excel",
                iconProps: {
                    iconName: "ExcelLogo"
                },
                onClick: onExportExcel
            });
        }
        return tmpCommandBarItemProps;
    }

    function getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, field: IField): IContextualMenuProps {
        const contextualMenuItems = [];
        const isFiltered = state.view.filteredBy.filter(filteredByItem => filteredByItem.property === field.fieldName).length > 0;
        const sortBy = state.view.sortedBy.filter(sortedByItem => sortedByItem.property === field.fieldName);
        const isGrouped = state.view.groupedBy.filter(groupedByItem => groupedByItem.fieldName === field.fieldName).length > 0;
        if (!props.disableSortBy) {
            contextualMenuItems.push({
                key: 'sortAsc',
                name: 'Ordre croissant',
                iconProps: { iconName: "SortUp" },
                canCheck: !(sortBy.length > 0 && !sortBy[0].isSortedDescending) ? true : false,
                checked: (sortBy.length > 0 && !sortBy[0].isSortedDescending) ? true : false,
                onClick: () => {
                    onSortField(field, false);
                }
            });
            contextualMenuItems.push({
                key: 'sortDesc',
                name: 'Ordre décroissant',
                iconProps: { iconName: "SortDown" },
                canCheck: !(sortBy.length > 0 && sortBy[0].isSortedDescending) ? true : false,
                checked: (sortBy.length > 0 && sortBy[0].isSortedDescending) ? true : false,
                onClick: () => {
                    onSortField(field, true);
                }
            });
        }
        if (!props.disableFilterBy) {
            contextualMenuItems.push({
                key: 'filterBy',
                name: 'Filtrer par ' + field.displayName,
                iconProps: { iconName: "Filter" },
                canCheck: true,
                checked: isFiltered,
                onClick: () => {
                    onFilterByField(field);
                }
            });
        }
        if (!props.disableGroupBy) {
            contextualMenuItems.push({
                key: 'groupBy',
                name: 'Regrouper par ' + field.displayName,
                iconProps: { iconName: "GroupedDescending" },
                canCheck: true,
                checked: isGrouped,
                onClick: () => {
                    onGroupByField(field);
                }
            });
        }
        return {
            items: contextualMenuItems,
            target: ev.currentTarget as HTMLElement,
            directionalHint: DirectionalHint.bottomAutoEdge,
            onDismiss: onContextualMenuDismissed
        };
    }

    function getExpandedGroups(groups: IGroup[]): string[] {
        let expandedGroups: string[] = [];
        if (groups) {
            groups.forEach(group => {
                if (!group.isCollapsed) {
                    expandedGroups.push(group.key);
                }
                if (group.children) {
                    expandedGroups.push(...getExpandedGroups(group.children));
                }
            });
        }
        return expandedGroups;
    }

    function getFilterPanelProps(field: IField): IFilterPanelProps {
        let defaultSelectedValues: string[] = [];
        const filteredBy: IFilterBy[] = [...state.view.filteredBy];
        for (let i = 0; i < filteredBy.length; i++) {
            if (filteredBy[i].property === field.fieldName) {
                defaultSelectedValues = filteredBy[i].values;
                break;
            }
        }
        return {
            isOpen: true,
            items: defaultSelectedValues.length > 0 ? props.items : state.items,
            property: field.fieldName,
            type: field.type,
            title: `Filtrer par « ${field.displayName} »`,
            defaultSelectedValues: defaultSelectedValues,
            onPropertyFilterChange: (selectedValues) => {
                onPropertyFilterChange(field.fieldName, selectedValues);
            },
            onDismissed: onFilterPanelDismissed
        } as IFilterPanelProps;
    }

    function getFormPanelProps(item: any = null): IFormPanelProps {
        return {
            panelProps: {
                isOpen: props.showFormPanel,
                isBlocking: true,
                type: props.formPanelType,
                closeButtonAriaLabel: "Fermer",
                onDismiss: () => {
                    setFormPanelProps(undefined);
                }
            },
            selectedItem: item
        };
    }

    function getGroups(view: IView, items: any[]): IGroup[] {
        const createGroup = (name: string, startIndex: number, count: number, level: number, isCollapsed: boolean): IGroup => {
            return {
                key: "group_" + level.toString() + "_" + name,
                name: name ? name : "",
                startIndex: startIndex,
                count: count,
                level: level,
                isCollapsed: isCollapsed
            };
        };

        const groupBy = (itemsToGroup: any[], groupedBy: IGroupBy[], sortedBy: ISortBy[], startIndex: number = 0, level: number = 0): IGroup[] => {
            let groups: IGroup[] = null;
            let groupName: string;
            let groupStartIndex: number;
            let counter: number;
            let group: IGroup;
            let propertyGroupBy: IGroupBy;
            let fieldGroupBy: IField;
            let tmpGroupedBy: IGroupBy[] = groupedBy.slice();
            let oldSortedBy: ISortBy[] = sortedBy.slice();
            let tmpSortedBy: ISortBy[] = [];
            let tmpItems: any[] = itemsToGroup.slice();
            if (tmpItems.length > 0 && tmpGroupedBy.length > 0) {
                if (level === 0) {
                    groupedBy.forEach(groupedByItem => {
                        let isSortedDescending = false;
                        for (let i = 0; i < sortedBy.length; i++) {
                            if (sortedBy[i].property === groupedByItem.fieldName) {
                                oldSortedBy.splice(i, 1);
                                isSortedDescending = sortedBy[i].isSortedDescending;
                                break;
                            }
                        }
                        tmpSortedBy.push({ property: groupedByItem.fieldName, isSortedDescending: isSortedDescending });
                    });
                    tmpSortedBy.push(...oldSortedBy);
                    tmpItems = M365ComponentsLibrary.sortBy(tmpItems, tmpSortedBy);
                }
                propertyGroupBy = tmpGroupedBy.shift();
                fieldGroupBy = view.fields.filter(field => field.fieldName === propertyGroupBy.fieldName)[0];
                if (propertyGroupBy) {
                    groups = [];
                    groupName = M365ComponentsLibrary.formatValueAsString(fieldGroupBy.type, M365ComponentsLibrary.getPropertyValueByPath(tmpItems[0], propertyGroupBy.fieldName));
                    groupStartIndex = startIndex;
                    counter = 0;
                    for (let i = 1; i < tmpItems.length; i++) {
                        counter++;
                        let itemPropertyValue: string = M365ComponentsLibrary.formatValueAsString(fieldGroupBy.type, M365ComponentsLibrary.getPropertyValueByPath(tmpItems[i], propertyGroupBy.fieldName));
                        if (groupName !== itemPropertyValue) {
                            group = createGroup(groupName, groupStartIndex, counter, level, propertyGroupBy.isCollapsed);
                            groupStartIndex = startIndex + i;
                            groupName = itemPropertyValue;
                            counter = 0;
                            if (group.count > 0) {
                                group.children = groupBy(tmpItems.slice(group.startIndex - startIndex, group.startIndex - startIndex + group.count), tmpGroupedBy, tmpSortedBy, group.startIndex, group.level + 1);
                                groups.push(group);
                            }
                        }
                    }
                    counter++;
                    group = createGroup(groupName, groupStartIndex, counter, level, propertyGroupBy.isCollapsed);
                    if (group.count > 0) {
                        group.children = groupBy(tmpItems.slice(group.startIndex - startIndex, group.startIndex - startIndex + group.count), tmpGroupedBy, tmpSortedBy, group.startIndex, group.level + 1);
                        groups.push(group);
                    }
                }
            }
            return groups;
        };

        return groupBy(items, view.groupedBy, view.sortedBy);
    }

    function getItems(view: IView, items: any[]): any[] {
        return M365ComponentsLibrary.sortBy(filter(items, view.filteredBy), view.sortedBy);
    }

    function init(initialProps: IAdvancedDetailsListProps): IAdvancedDetailsListState {
        const newItems = getItems(initialProps.defaultView, initialProps.items);
        return {
            items: newItems,
            groups: getGroups(initialProps.defaultView, newItems),
            columns: getColumns(initialProps.defaultView),
            view: initialProps.defaultView
        };
    }

    function getKey(item: any, index?: number): string {
        if (!props.itemKeyProperty) {
            return index.toString();
        } else {
            return M365ComponentsLibrary.getPropertyValueByPath(item, props.itemKeyProperty);
        }
    }

    function onSortField(field: IField, isSortedDescending: boolean): void {
        state.view.sortedBy = [{ property: field.fieldName, isSortedDescending: isSortedDescending }];
        dispatch(state.view);
    }

    function onContextualMenuDismissed(): void {
        setContextualMenuProps(undefined);
    }

    function onEdit() {
        const selectedItem: any = selection.getSelectedCount() > 0 ? selection.getSelection()[0] : null;
        if (props.onEdit) {
            props.onEdit(selectedItem);
        } else {
            if (props.onRenderFormPanel) {
                setFormPanelProps(getFormPanelProps(selectedItem));
            }
        }
    }

    function onExportExcel(): void {
        const items = state.items;
        const fields = state.view.fields;
        const filename = "export.xlsx";
        let grid: any[] = new Array(items.length);
        let headers: string[] = [];
        fields.forEach(field => {
            headers.push(field.displayName);
        });
        grid[0] = headers;
        items.forEach((item: any, index: number) => {
            let row: string[] = [];
            fields.forEach(field => {
                row.push(M365ComponentsLibrary.formatValueAsString(field.type, M365ComponentsLibrary.getPropertyValueByPath(item, field.fieldName)));
            });
            grid[index + 1] = row;
        });
        let ws = xlsx.utils.aoa_to_sheet(grid);
        ws["!autofilter"] = { ref: "A1:S1" };
        let wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, ws);
        let wbout = xlsx.write(wb, { bookType: 'xlsx', bookSST: false, type: 'array' });
        saveAs(new Blob([wbout], { type: "application/octet-stream" }), filename);
    }

    function onFilterByField(field: IField): void {
        setFilterPanelProps(getFilterPanelProps(field));
    }

    function onFilterPanelDismissed(): void {
        setFilterPanelProps(undefined);
    }

    function onGroupByField(field: IField): void {
        let groupedBy = [...state.view.groupedBy];
        let index = -1;
        for (let i = 0; i < groupedBy.length; i++) {
            if (groupedBy[i].fieldName === field.fieldName) {
                index = i;
                break;
            }
        }
        if (index != -1) {
            groupedBy.splice(index, 1);
        } else {
            groupedBy.push({ fieldName: field.fieldName, isCollapsed: true });
        }
        state.view.groupedBy = groupedBy;
        dispatch(state.view);
    }

    function onNew(): void {
        if (props.onNew) {
            props.onNew();
        } else {
            if (props.onRenderFormPanel) {
                setFormPanelProps(getFormPanelProps());
            }
        }
    }

    function onRemove(): void {
        const selectedItems: any = selection.getSelection();
        if (props.onRemove) {
            if (props.onRemoveItems) {
                props.onRemoveItems(selectedItems);
            }
        }
    }

    function onPropertyFilterChange(columnKey: string, filterValues: string[]): void {
        let filteredBy = [...state.view.filteredBy];
        let exists = false;
        let index = 0;
        for (let i = 0; i < filteredBy.length; i++) {
            if (filteredBy[i].property === columnKey) {
                index = i;
                exists = true;
                break;
            }
        }
        if (filterValues && filterValues.length > 0) {
            if (!exists) {
                filteredBy.push({ property: columnKey, values: filterValues });
            } else {
                filteredBy[index].values = filterValues;
            }
        } else {
            filteredBy.splice(index, 1);
        }
        state.view.filteredBy = filteredBy;
        dispatch(state.view);
    }

    function reducer(oldState: IAdvancedDetailsListState, nextView: IView): IAdvancedDetailsListState {
        let oldSortedBy: ISortBy[] = nextView.sortedBy.slice();
        let tmpSortedBy: ISortBy[] = [];
        nextView.groupedBy.forEach(groupBy => {
            let isSortedDescending = false;
            for (let i = 0; i < nextView.sortedBy.length; i++) {
                if (nextView.sortedBy[i].property === groupBy.fieldName) {
                    oldSortedBy.splice(i, 1);
                    isSortedDescending = nextView.sortedBy[i].isSortedDescending;
                    break;
                }
            }
            tmpSortedBy.push({ property: groupBy.fieldName, isSortedDescending: isSortedDescending });
        });
        tmpSortedBy.push(...oldSortedBy);
        nextView.sortedBy = tmpSortedBy;
        let newItems: any[] = getItems(nextView, props.items);
        let expandedGroups: string[] = getExpandedGroups(oldState.groups);
        return {
            items: newItems,
            groups: setExpandedGroups(getGroups(nextView, newItems), expandedGroups),
            columns: getColumns(nextView),
            view: nextView
        };
    }

    function setExpandedGroups(groups: IGroup[], expandedGroups: string[]): IGroup[] {
        let tmpGroups = null;
        if (groups) {
            tmpGroups = [...groups];
            if (tmpGroups) {
                tmpGroups.forEach(group => {
                    if (expandedGroups.indexOf(group.key) != -1) {
                        group.isCollapsed = false;
                    } else {
                        group.isCollapsed = true;
                    }
                    if (group.children) {
                        group.children = setExpandedGroups(group.children, expandedGroups);
                    }
                });
            }
        }
        return tmpGroups;
    }

    return (
        <div>
            {
                !props.disableCommandBar &&
                commandBarItemProps &&
                <CommandBar items={commandBarItemProps} />
            }
            <DetailsList
                selection={selection}
                items={state.items}
                columns={state.columns}
                groups={state.groups}
                getKey={getKey}
                setKey="multiple"
                layoutMode={DetailsListLayoutMode.justified}
                className={props.className}
                compact={props.compact}
                onItemInvoked={props.onItemInvoked}
                selectionMode={props.selectionMode}
            />
            {
                formPanelProps &&
                formPanelProps.panelProps &&
                <Panel {...formPanelProps.panelProps}>
                    {
                        props.onRenderFormPanel &&
                        props.onRenderFormPanel(
                            { ...formPanelProps.selectedItem },
                            () => {
                                setFormPanelProps(undefined);
                            }
                        )
                    }
                </Panel>
            }
            {contextualMenuProps && <ContextualMenu {...contextualMenuProps} />}
            {filterPanelProps && <FilterPanel {...filterPanelProps} />}
        </div>
    );
};

export { AdvancedDetailsList };
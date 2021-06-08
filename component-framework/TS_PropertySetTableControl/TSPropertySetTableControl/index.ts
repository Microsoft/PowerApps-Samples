/*
	This file is part of the Microsoft PowerApps code samples. 
	Copyright (C) Microsoft Corporation.  All rights reserved. 
	This source code is intended only as a supplement to Microsoft Development Tools and/or  
	on-line documentation.  See these other materials for detailed information regarding  
	Microsoft code samples. 

	THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER  
	EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF  
	MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
 */


import {IInputs, IOutputs} from "./generated/ManifestTypes";

import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { table } from "console";
type DataSet = ComponentFramework.PropertyTypes.DataSet;


	// Define const here
	const RowRecordId:string = "rowRecId";

	// Style name of disabled buttons
	const Button_Disabled_style =  "loadNextPageButton_Disabled_Style";

	export class TSPropertySetTableControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

		private contextObj: ComponentFramework.Context<IInputs>;
		
		// Div element created as part of this control's main container
		private mainContainer: HTMLDivElement;

		// Table element created as part of this control's table
		private dataTable: HTMLTableElement;

		private getValueResultDiv: HTMLDivElement;

		private pagingDiv: HTMLDivElement;

		private selectedRecord: DataSetInterfaces.EntityRecord;

		private selectedRecords: {[id: string]: boolean} = {};

		private displayPropertySetColumns = true;

		private targetEntityDiv: HTMLDivElement;

		private sortedColumn = "";

		private direction: DataSetInterfaces.Types.SortDirection = 0;

		private logs: string[] = [];
		/**
		 * Empty constructor.
		 */
		constructor()
		{
		}

		/**
		 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
		 * Data-set values are not initialized here, use updateView.
		 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
		 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
		 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
		 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
		 */
		public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
		{
			// Need to track container resize so that control could get the available width.
			// In Model-driven app, the available height won't be provided even this is true
			// In Canvas-app, the available height will be provided in context.mode.allocatedHeight
			context.mode.trackContainerResize(true);

			// Create main table container div.
			this.mainContainer = document.createElement("div");
			this.mainContainer.classList.add("SimpleTable_MainContainer_Style");
			this.mainContainer.id = "SimpleTableMainContainer";
			// Create data table container div.
			this.dataTable = document.createElement("table");
			this.dataTable.classList.add("SimpleTable_Table_Style");
			this.pagingDiv = this.createPagingDiv(context);
			this.targetEntityDiv = document.createElement("div");
			this.targetEntityDiv.classList.add("StatusDiv＿Style");

			// Create main table container div. 
			this.mainContainer = document.createElement("div");

			// Adding the main table and loadNextPage button created to the container DIV.
			this.mainContainer.appendChild(this.createSearchBar(context));
			//this.mainContainer.appendChild(refreshTestButton);
			this.mainContainer.appendChild(this.dataTable);
			this.mainContainer.appendChild(this.pagingDiv);
			this.mainContainer.classList.add("main-container");
			container.appendChild(this.mainContainer);
		}

	private _refresh() {
		this.contextObj.parameters.sampleDataSet.refresh();
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		this.contextObj = context;
		const param = context.parameters.sampleDataSet;
		this.targetEntityDiv.innerText = `${JSON.stringify(context.updatedProperties)} | ${param.getTargetEntityType()} | ${param.loading ? "loading" : "done"} | ${param.getTitle() ? param.getTitle() + "(" + param.getViewId() + ")" : ""})`;
		if (!context.parameters.sampleDataSet.loading) {
			// Get sorted columns on View
			let columnsOnView = this.getSortedColumnsOnView(context);
			if (!columnsOnView || columnsOnView.length === 0) {
				return;
			}

			//calculate the width for each column
			let columnWidthDistribution = this.getColumnWidthDistribution(context, columnsOnView);

			//When new data is received, it needs to first remove the table element, allowing it to properly render a table with updated data
			//This only needs to be done on elements having child elements which is tied to data received from canvas/model ..
			while (this.dataTable.firstChild) {
				this.dataTable.removeChild(this.dataTable.firstChild);
			}

			this.dataTable.appendChild(this.createTableHeader(columnsOnView, columnWidthDistribution));
			this.dataTable.appendChild(this.createTableBody(columnsOnView, columnWidthDistribution, context.parameters.sampleDataSet));
			this.dataTable.style.height = (context.mode.allocatedHeight - 160) + "px";
		}
		this.pagingDiv.remove();
		this.pagingDiv = this.createPagingDiv(context);
		this.mainContainer.appendChild(this.pagingDiv);
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/**
		 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
	}

	private createGetValueDiv(): HTMLDivElement {
		const getValueDiv = document.createElement("div");
		const inputBox = document.createElement("input");
		const getValueButton = document.createElement("button");
		const addColumnButton = document.createElement("button");
		const resultDiv = document.createElement("div");
		const _this = this; 

		inputBox.id = "getValueInputBox";
		inputBox.placeholder = "select a row and enter the alias name";
		inputBox.classList.add("GetValueInput_Style");

		getValueButton.innerText = "GetValue";		
		getValueButton.classList.add("Button_Style");
		getValueButton.onclick = () => {
			if (_this.selectedRecord) {
				if (this.getValueResultDiv) {
					this.getValueResultDiv.innerHTML = "";
				}
				const alias = inputBox.value;
				const value = _this.selectedRecord.getValue(alias);
				const formattedValue = _this.selectedRecord.getFormattedValue(alias);
				const recordId = _this.selectedRecord.getRecordId();
				const namedReference = _this.selectedRecord.getNamedReference();
				const content1 = document.createElement("div");
				const content2 = document.createElement("div");
				const content3 = document.createElement("div");
				const content4 = document.createElement("div");
				content1.innerText= `Value: ${value}`;
				content2.innerText= `FormattedValue: ${formattedValue}`;
				content3.innerText= `RecordId: ${recordId}`;
				content4.innerText= `NamedReference: ${JSON.stringify(namedReference)}`;
				resultDiv.appendChild(content1);
				resultDiv.appendChild(content2);
				resultDiv.appendChild(content3);
				resultDiv.appendChild(content4);
			}
		};

		addColumnButton.innerText = "AddColumn";
		addColumnButton.classList.add("Button_Style");
		addColumnButton.onclick = () => {
			if (inputBox.value) {
				_this.contextObj.parameters.sampleDataSet.addColumn?.(inputBox.value);
				_this._refresh();
			}
		};

		resultDiv.classList.add("GetValueResult_Style");
		resultDiv.innerText = "Select a row first";
		this.getValueResultDiv = resultDiv;
		getValueDiv.appendChild(inputBox);
		getValueDiv.appendChild(getValueButton);
		getValueDiv.appendChild(addColumnButton);
		getValueDiv.appendChild(resultDiv);
		//getValueDiv.appendChild(newValueInput);
		//getValueDiv.appendChild(saveValueButton);
		return getValueDiv;
	}

	/**
	 * Get sorted columns on view, columns are sorted by DataSetInterfaces.Column.order
	 * Property-set columns will always have order = -1.
	 * In Model-driven app, the columns are ordered in the same way as columns defined in views.
	 * In Canvas-app, the columns are ordered by the sequence fields added to control
	 * Note that property set columns will have order = 0 in test harness, this is a bug.
	 * @param context
	 * @return sorted columns object on View
	 */
	private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[] {
		if (!context.parameters.sampleDataSet.columns) {
			return [];
		}

		let columns = context.parameters.sampleDataSet.columns;

		return columns;
	}

	/**
	 * Get column width distribution using visualSizeFactor. 
	 * In model-driven app, visualSizeFactor can be configured from view's settiong.
	 * In Canvas app, currently there is no way to configure this value. In all data sources, all columns will have the same visualSizeFactor value.
	 * Control does not have to render the control using these values, controls are free to display any columns with any width, or making column width adjustable.
	 * However, these kind of configurations will be lost when leaving the page
	 * @param context context object of this cycle
	 * @param columnsOnView columns array on the configured view
	 * @returns column width distribution
	 */
	private getColumnWidthDistribution(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]): string[] {

		let widthDistribution: string[] = [];

		// Considering need to remove border & padding length
		let totalWidth: number = context.mode.allocatedWidth - 60;
		let widthSum = 0;
		let defaultSortColumn: string = "";
		columnsOnView.forEach(function (columnItem) {
			widthSum += columnItem.visualSizeFactor;
			defaultSortColumn = defaultSortColumn || columnItem.name;
		});
		this.sortedColumn = this.sortedColumn || defaultSortColumn;

		let remainWidth: number = totalWidth;

		columnsOnView.forEach(function (item, index) {
			let widthPerCell = "";
			if (index !== columnsOnView.length - 1) {
				let cellWidth = Math.floor((item.visualSizeFactor / widthSum) * totalWidth);
				remainWidth = remainWidth - cellWidth;
				widthPerCell = cellWidth + "px";
			}
			else {
				widthPerCell = remainWidth + "px";
			}
			widthDistribution.push(widthPerCell);
		});

		return widthDistribution;

	}

	private createTableHeader(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[]): HTMLTableSectionElement {
		let tableHeader: HTMLTableSectionElement = document.createElement("thead");
		let tableHeaderRow: HTMLTableRowElement = document.createElement("tr");
		tableHeaderRow.classList.add("SimpleTable_TableRow_Style");
		const displayPropertySetColumns = this.displayPropertySetColumns;
		const _this = this;
		const headerSelectedColumn = document.createElement("th");
		headerSelectedColumn.width = "60px";
		headerSelectedColumn.innerText = "Select";
		headerSelectedColumn.classList.add("SimpleTable_TableHeader_Selected_Style");
		tableHeaderRow.appendChild(headerSelectedColumn);
			
		columnsOnView.forEach(function (columnItem, index) {
			if (columnItem.order >= 0 || displayPropertySetColumns) {
				let tableHeaderCell = document.createElement("th");
				let innerDiv = document.createElement("div");
				tableHeaderCell.width = widthDistribution[index];
				innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
				innerDiv.style.maxWidth = widthDistribution[index];
				let columnDisplayName: string;
				if (columnItem.order < 0) {
					tableHeaderCell.classList.add("SimpleTable_TableHeader_PropertySet_Style");
					columnDisplayName = columnItem.displayName + "(propertySet)";
				} else {
					tableHeaderCell.classList.add("SimpleTable_TableHeader_Style");
					columnDisplayName = columnItem.displayName;
				}
				if (columnItem.name === _this.sortedColumn) {
					columnDisplayName += _this.direction === 1 ? " ↓" : " ↑";
				}
				innerDiv.innerText = columnDisplayName;

				tableHeaderCell.appendChild(innerDiv);
				tableHeaderCell.addEventListener("click", (() => {
					if (_this.sortedColumn !== columnItem.name) {
						_this.sortedColumn = columnItem.name;
					} else {
						_this.direction = (1 - _this.direction) as DataSetInterfaces.Types.SortDirection;
					}
					_this.contextObj.parameters.sampleDataSet.refresh();
				}).bind(_this));
				tableHeaderRow.appendChild(tableHeaderCell);
			}
		});
		tableHeader.appendChild(tableHeaderRow);
		return tableHeader;
	}

	private createTableBody(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], gridParam: DataSet): HTMLTableSectionElement {
		const tableBody: HTMLTableSectionElement = document.createElement("tbody");
		const displayPropertySetColumns = this.displayPropertySetColumns;
		const selectedRecordIds = this.contextObj.parameters.sampleDataSet.getSelectedRecordIds();
		const _this = this;
		if (gridParam.sortedRecordIds.length > 0) {
			for (let currentRecordId of gridParam.sortedRecordIds) {

				let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
				tableRecordRow.classList.add("SimpleTable_TableRow_Style");
				tableRecordRow.addEventListener("click", this.onRowClick.bind(this));

				// Set the recordId on the row dom, this is the simplest way to help us track which record has been clicked.
				tableRecordRow.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());
				const selected = selectedRecordIds?.indexOf(currentRecordId) > -1;
				let tableSelectedCell = document.createElement("td");
				tableSelectedCell.width = "64px";
				if (selected) {
					tableSelectedCell.innerText = "+";
				}
				tableRecordRow.appendChild(tableSelectedCell);
				columnsOnView.forEach(function (columnItem, index) {
					if (columnItem.order >= 0 || displayPropertySetColumns) {
						let tableRecordCell = document.createElement("td");
						tableRecordCell.classList.add("SimpleTable_TableCell_Style");
						let innerLink = document.createElement("a");
						innerLink.innerText = gridParam.records[currentRecordId].getFormattedValue(columnItem.alias);
						let innerDiv = document.createElement("div");
						innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
						innerDiv.style.width = widthDistribution[index];
						innerDiv.appendChild(innerLink);
						tableRecordCell.appendChild(innerDiv);
						tableRecordRow.appendChild(tableRecordCell);
						innerLink.addEventListener("click", _this.onLinkClick.bind(_this));
					}
				});

				tableBody.appendChild(tableRecordRow);
			}
		}
		else {
			let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
			let tableRecordCell: HTMLTableCellElement = document.createElement("td");
			tableRecordCell.classList.add("No_Record_Style");
			tableRecordCell.colSpan = columnsOnView.length;
			tableRecordCell.innerText = this.contextObj.resources.getString("TSPropertySetTableControl_No_Record_Found");
			tableRecordRow.appendChild(tableRecordCell)
			tableBody.appendChild(tableRecordRow);
		}

		return tableBody;
	}


	/**
	 * Row Click Event handler for the associated row when being clicked
	 * @param event
	 */
	private onRowClick(event: Event): void {
		let rowElement = (event.currentTarget as HTMLTableRowElement);
		let rowRecordId = rowElement.getAttribute(RowRecordId);
		if (rowRecordId) {
			const record = this.contextObj.parameters.sampleDataSet.records[rowRecordId];
			this.selectedRecord = record;
			if (this.getValueResultDiv) {
				this.getValueResultDiv.innerHTML = "";
			}
			this.selectedRecords[rowRecordId] = !this.selectedRecords[rowRecordId];
			const selectedRecordsArray = [];
			for (const recordId in this.selectedRecords) {
				if (this.selectedRecords[recordId]) {
					selectedRecordsArray.push(recordId);
				}
			}
			this.contextObj.parameters.sampleDataSet.setSelectedRecordIds(selectedRecordsArray);
			this.contextObj.factory.requestRender();
		}
	}

	private onLinkClick(event: Event): void {
		let rowElement = (event.currentTarget as HTMLTableRowElement);
		let rowRecordId = rowElement.getAttribute(RowRecordId);
		if (rowRecordId) {
			const record = this.contextObj.parameters.sampleDataSet.records[rowRecordId];
			this.selectedRecord = record;
			if (this.getValueResultDiv) {
				this.getValueResultDiv.innerHTML = "";
			}
			this.selectedRecords[rowRecordId] = !this.selectedRecords[rowRecordId];
			const selectedRecordsArray = [];
			for (const recordId in this.selectedRecords) {
				if (this.selectedRecords[recordId]) {
					selectedRecordsArray.push(recordId);
				}
			}
			this.contextObj.parameters.sampleDataSet.openDatasetItem(record.getNamedReference());
		}
	}

	/**
	 * 'Load Next' Button Event handler when load more button clicks
	 * @param event
	 */
	private onLoadNextButtonClick(event: Event): void {
		this.contextObj.parameters.sampleDataSet.paging.loadNextPage();
	}

	/**
	 * 'Load Prevous' Button Event handler when load more button clicks
	 * @param event
	 */
	private onLoadPrevButtonClick(event: Event): void {
		this.contextObj.parameters.sampleDataSet.paging.loadPreviousPage();
	}

	private createPagingDiv(context: ComponentFramework.Context<IInputs>) {
		let container = document.createElement("div");
		let loadPrevPageButton: HTMLButtonElement;
		let loadNextPageButton: HTMLButtonElement;
		let setPageSizeButton: HTMLButtonElement;
		let setPageInput: HTMLInputElement;
		loadPrevPageButton = document.createElement("button");
		loadPrevPageButton.setAttribute("type", "button");
		loadPrevPageButton.innerText = context.resources.getString("TSPropertySetTableControl_LoadPrev_ButtonLabel");
		loadPrevPageButton.classList.add("Button_Style");
		loadPrevPageButton.addEventListener("click", this.onLoadPrevButtonClick.bind(this));
		loadPrevPageButton.disabled = !context.parameters.sampleDataSet.paging.hasPreviousPage;
		if (!context.parameters.sampleDataSet.paging.hasPreviousPage) {
			loadPrevPageButton.classList.add(Button_Disabled_style);
		}

		loadNextPageButton = document.createElement("button");
		loadNextPageButton.setAttribute("type", "button");
		loadNextPageButton.innerText = context.resources.getString("TSPropertySetTableControl_LoadNext_ButtonLabel");
		loadNextPageButton.classList.add(Button_Disabled_style);
		loadNextPageButton.classList.add("Button_Style");
		loadNextPageButton.addEventListener("click", this.onLoadNextButtonClick.bind(this));
		loadNextPageButton.disabled = !context.parameters.sampleDataSet.paging.hasNextPage;
		if (!context.parameters.sampleDataSet.paging.hasNextPage) {
			loadNextPageButton.classList.add(Button_Disabled_style);
		}
		setPageInput = document.createElement("input");
		setPageInput.classList.add("SetPageSizeInput_style");
		setPageInput.id = "setPageInput";
		setPageSizeButton = document.createElement("button");
		setPageSizeButton.setAttribute("type", "button");
		setPageSizeButton.innerText = context.resources.getString("TSPropertySetTableControl_SetPageSize_ButtonLabel");
		setPageSizeButton.classList.add("Button_Style");
		setPageSizeButton.addEventListener("click", () => {
			const pageSize = parseInt(setPageInput.value || "25", 10);
			context.parameters.sampleDataSet.paging.setPageSize(pageSize);
			context.parameters.sampleDataSet.refresh();
		});

		
		container.appendChild(loadPrevPageButton);
		container.appendChild(loadNextPageButton);
		container.appendChild(setPageInput);
		container.appendChild(setPageSizeButton);

		return container;
	}

	private createSearchBar(context: ComponentFramework.Context<IInputs>) {
		const columns = this.getSortedColumnsOnView(context);
		let container = document.createElement("div");
		let input = document.createElement("input");
		input.placeholder = "Search records";
		input.classList.add("GetValueInput_Style");
		input.id = "searchBar";
		let button = document.createElement("button");
		button.classList.add("Button_Style");
		button.innerHTML = "Search"
		button.addEventListener("click", (() => {
			let conditionsArray: DataSetInterfaces.ConditionExpression[] = [];
			let searchString = input.value;
			for (let i = 0; i < columns.length; i++) {
				const column = columns[i];
				if (!column.isHidden && column.dataType === "SingleLine.Text" && column.name) {
					const condition: DataSetInterfaces.ConditionExpression = {
						attributeName: column.name,
						conditionOperator: 6,
						value: searchString + "%",
					};
					conditionsArray.push(condition);
				}
			}
			this.contextObj.parameters.sampleDataSet.filtering.setFilter({
				conditions: conditionsArray,
				filterOperator: 1,
			});
			if (this.sortedColumn) {
				this.contextObj.parameters.sampleDataSet.sorting = [{
					sortDirection: this.direction,
					name: this.sortedColumn,
				}];
			}
			
			this.contextObj.parameters.sampleDataSet.refresh();
		}).bind(this));
		const firstDiv = document.createElement("div");
		firstDiv.appendChild(input);
		firstDiv.appendChild(button);
		firstDiv.appendChild(this.targetEntityDiv);
		container.appendChild(firstDiv);
		container.appendChild(this.createGetValueDiv());	
		return container;
	}
}
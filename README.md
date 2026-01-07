<?xml version="1.0" encoding="utf-8" ?>
<manifest>
  <control namespace="DonSchia" constructor="FetchXmlDetailsList" version="3.0.0" display-name-key="FetchXml DetailsList" description-key="FluentUI DetailsList for views, subgrids and forms with custom FetchXml and pagination" control-type="virtual">
    
    <!-- Dataset binding for views/grids -->
    <data-set name="dataSetGrid" display-name-key="Grid Data" cds-data-set-options="displayCommandBar:true;displayViewSelector:true;displayQuickFindSearch:true">
      <property-set name="FetchXml" display-name-key="Custom FetchXml" description-key="Optional custom FetchXml query (leave empty to use view's query)" of-type="Multiple" usage="input" required="false" />
      <property-set name="ColumnLayoutJson" display-name-key="Column Layout JSON" description-key="Optional JSON column configuration (leave empty to use view's columns)" of-type="Multiple" usage="input" required="false" />
    </data-set>
    
    <!-- Global Properties -->
    <property name="ItemsPerPage" display-name-key="Items Per Page" description-key="Records per page (default: 50)" of-type="Whole.None" usage="input" required="false" />
    
    <property name="DebugMode" display-name-key="Debug Mode" description-key="Enable debug logging" of-type="Enum" usage="input" required="false">
      <value name="Off" display-name-key="Off">0</value>
      <value name="On" display-name-key="On">1</value>
    </property>
    
    <!-- Resources -->
    <resources>
      <code path="index.ts" order="1"/>
      <css path="css/FetchXmlDetailsList.css" order="1" />
      <resx path="strings/FetchXmlDetailsList.1033.resx" version="1.0.0" />
    </resources>
    
    <!-- Feature usage -->
    <feature-usage>
      <uses-feature name="WebAPI" required="true" />
      <uses-feature name="Utility" required="false" />
    </feature-usage>
  </control>
</manifest>


import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { DynamicDetailsList } from "./DynamicsDetailsList";
import { GetSampleData } from "./GetSampleData";
import * as React from "react";
import * as ReactDOM from "react-dom";

export class FetchXmlDetailsList implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _isTestHarness: boolean = false;
    private _currentPage: number = 1;
    private _pageSize: number = 50;
    private _totalRecords: number = 0;
    private _dataset: ComponentFramework.PropertyTypes.DataSet | null = null;
    
    constructor() {
        // Empty constructor
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        try {
            this._context = context;
            this._notifyOutputChanged = notifyOutputChanged;
            this._container = container;
            
            if (this._container) {
                this._container.style.width = "100%";
                this._container.style.height = "100%";
            }
            
            // Check if we have dataset (view mode)
            const params = context.parameters as any;
            this._dataset = params.dataSetGrid;
            
            this._isTestHarness = this.isTestHarness();
            
            const itemsPerPage = params.ItemsPerPage?.raw;
            this._pageSize = itemsPerPage ? itemsPerPage : 50;
            
            console.log("FetchXmlDetailsList initialized", {
                isTestHarness: this._isTestHarness,
                pageSize: this._pageSize,
                hasDataset: !!this._dataset,
                mode: this._dataset ? "VIEW/GRID" : "TEST HARNESS"
            });
        } catch (error) {
            console.error("Error in init:", error);
            this.showError("Initialization failed: " + (error as Error).message);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        try {
            this._context = context;
            const params = context.parameters as any;
            this._dataset = params.dataSetGrid;
            
            // Dataset mode (Views/Grids in Dynamics)
            if (this._dataset && this._dataset.loading === false) {
                console.log("Loading from dataset (VIEW MODE)");
                this.loadFromDataset();
            }
            // Test harness mode (npm start)
            else if (this.isTestHarness()) {
                console.log("Loading sample data (TEST HARNESS)");
                this.loadSampleData();
            }
            // Form mode with custom FetchXml (fallback)
            else {
                console.log("Dataset still loading or not available");
                this.showLoading();
            }
        } catch (error) {
            console.error("Error updating view:", error);
            this.showError("Failed to load data: " + (error as Error).message);
        }
    }

    private isTestHarness(): boolean {
        try {
            const webAPI = (this._context as any)?.webAPI;
            const hasDataset = !!this._dataset;
            // If no dataset and no webAPI, it's test harness
            return !hasDataset && (!webAPI || typeof webAPI.retrieveMultipleRecords !== 'function');
        } catch {
            return true;
        }
    }

    private loadFromDataset(): void {
        try {
            if (!this._dataset) {
                console.error("Dataset is null");
                return;
            }

            const debugMode = this.getDebugMode();
            
            if (debugMode) {
                console.log("Dataset:", this._dataset);
                console.log("Dataset columns:", this._dataset.columns);
                console.log("Dataset records:", this._dataset.sortedRecordIds?.length || 0);
            }

            // Check for custom FetchXml
            const customFetchXml = (this._dataset as any).FetchXml?.raw;
            const customColumnLayout = (this._dataset as any).ColumnLayoutJson?.raw;

            // Build items from dataset
            const items: any[] = [];
            if (this._dataset.sortedRecordIds) {
                this._dataset.sortedRecordIds.forEach((recordId) => {
                    const record = this._dataset!.records[recordId];
                    if (record) {
                        const item: any = {};
                        
                        // Get all column values
                        this._dataset!.columns.forEach((column) => {
                            const value = record.getValue(column.name);
                            item[column.name] = value;
                            
                            // Also get formatted value if available
                            const formatted = record.getFormattedValue(column.name);
                            if (formatted) {
                                item[column.name + "@OData.Community.Display.V1.FormattedValue"] = formatted;
                            }
                        });
                        
                        // Add record ID
                        item[this._dataset!.getTargetEntityType() + "id"] = record.getRecordId();
                        items.push(item);
                    }
                });
            }

            // Build columns
            let columns: any[];
            
            if (customColumnLayout) {
                // Use custom column layout if provided
                try {
                    columns = JSON.parse(customColumnLayout);
                } catch (e) {
                    console.warn("Failed to parse custom column layout, using dataset columns");
                    columns = this.buildColumnsFromDataset();
                }
            } else {
                // Use dataset columns
                columns = this.buildColumnsFromDataset();
            }

            // Get pagination info
            const paging = this._dataset.paging;
            this._totalRecords = paging.totalResultCount || items.length;

            if (debugMode) {
                console.log("Items:", items);
                console.log("Columns:", columns);
                console.log("Total records:", this._totalRecords);
                console.log("Has next page:", paging.hasNextPage);
                console.log("Has previous page:", paging.hasPreviousPage);
            }

            const props: any = {
                context: this._context,
                items: items,
                columns: columns,
                primaryEntityName: this._dataset.getTargetEntityType(),
                isDebugMode: debugMode,
                baseD365Url: this.getBaseEnvironmentUrl(),
                pagination: {
                    currentPage: this._currentPage,
                    pageSize: this._pageSize,
                    totalRecords: this._totalRecords,
                    onPageChange: (page: number) => this.handleDatasetPageChange(page)
                }
            };

            this.renderDetailsList(props);
        } catch (error) {
            console.error("Error loading from dataset:", error);
            this.showError("Failed to load dataset: " + (error as Error).message);
        }
    }

    private buildColumnsFromDataset(): any[] {
        if (!this._dataset) return [];
        
        return this._dataset.columns.map((col) => ({
            key: col.name,
            fieldName: col.name,
            name: col.displayName,
            minWidth: col.visualSizeFactor || 100,
            maxWidth: 300,
            isResizable: true,
            data: {}
        }));
    }

    private handleDatasetPageChange(page: number): void {
        if (!this._dataset || !this._dataset.paging) return;
        
        const paging = this._dataset.paging;
        
        if (page > this._currentPage && paging.hasNextPage) {
            paging.loadNextPage();
            this._currentPage = page;
        } else if (page < this._currentPage && paging.hasPreviousPage) {
            paging.loadPreviousPage();
            this._currentPage = page;
        } else if (page === 1 && paging.hasPreviousPage) {
            // Go to first page
            paging.reset();
            this._currentPage = 1;
        }
    }

    private loadSampleData(): void {
        try {
            console.log("Loading sample data for test harness");
            
            const sampleData = GetSampleData();
            const startIndex = (this._currentPage - 1) * this._pageSize;
            const endIndex = startIndex + this._pageSize;
            const paginatedItems = sampleData.dataItems.slice(startIndex, endIndex);
            
            this._totalRecords = sampleData.dataItems.length;
            
            const props: any = {
                context: this._context,
                items: paginatedItems,
                columns: sampleData.columns,
                primaryEntityName: sampleData.primaryEntityName,
                isDebugMode: this.getDebugMode(),
                baseD365Url: this.getBaseEnvironmentUrl(),
                pagination: {
                    currentPage: this._currentPage,
                    pageSize: this._pageSize,
                    totalRecords: this._totalRecords,
                    onPageChange: (page: number) => this.handlePageChange(page)
                }
            };
            
            this.renderDetailsList(props);
        } catch (error) {
            console.error("Error loading sample data:", error);
            this.showError("Failed to load sample data: " + (error as Error).message);
        }
    }

    private getDebugMode(): boolean {
        try {
            const params = this._context?.parameters as any;
            const debugValue = params?.DebugMode?.raw;
            return debugValue === "1";
        } catch {
            return false;
        }
    }

    private getBaseEnvironmentUrl(): string {
        try {
            const pageContext = (this._context as any)?.page;
            
            if (pageContext && typeof pageContext.getClientUrl === 'function') {
                try {
                    return pageContext.getClientUrl();
                } catch (e) {
                    console.debug("getClientUrl failed");
                }
            }
            
            if (typeof window !== 'undefined' && window.location) {
                return window.location.origin;
            }
            
            return "";
        } catch (error) {
            console.debug("Could not get base environment URL:", error);
            return typeof window !== 'undefined' && window.location ? window.location.origin : "";
        }
    }

    private handlePageChange(page: number): void {
        this._currentPage = page;
        this.updateView(this._context);
    }

    private renderDetailsList(props: any): void {
        try {
            if (!this._container) {
                console.error("Container not initialized");
                return;
            }

            while (this._container.firstChild) {
                this._container.removeChild(this._container.firstChild);
            }

            const element = React.createElement(DynamicDetailsList, props);
            ReactDOM.render(element, this._container);
        } catch (error) {
            console.error("Error rendering details list:", error);
            this.showError("Failed to render: " + (error as Error).message);
        }
    }

    private showLoading(): void {
        if (!this._container) return;
        
        this._container.innerHTML = `
            <div style="display: flex; align-items: center; justify-content: center; height: 200px;">
                <div style="text-align: center;">
                    <div style="font-size: 24px; margin-bottom: 10px;">⏳</div>
                    <p style="color: #666;">Loading data...</p>
                </div>
            </div>
        `;
    }

    private showError(message: string): void {
        if (!this._container) return;
        
        this._container.innerHTML = `
            <div style="padding: 20px; background-color: #fef0f0; border: 1px solid #f5c6cb; border-radius: 4px;">
                <strong style="color: #a94442;">Error:</strong>
                <p style="color: #a94442; margin: 10px 0 0 0;">${message}</p>
            </div>
        `;
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        try {
            if (this._container) {
                ReactDOM.unmountComponentAtNode(this._container);
            }
        } catch (error) {
            console.error("Error in destroy:", error);
        }
    }
}

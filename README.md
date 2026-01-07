# PCF_TooltipControl

<?xml version="1.0" encoding="utf-8" ?>
<manifest>
  <control namespace="DonSchia" constructor="FetchXmlDetailsList" version="2.0.0" display-name-key="FetchXml DetailsList" description-key="FluentUI DetailsList for views and subgrids with FetchXml and pagination" control-type="virtual">
    
    <!-- Dataset binding - this allows the control to work on views/subgrids -->
    <data-set name="dataSet" display-name-key="Dataset" description-key="Dataset to display in the grid">
      <property-set name="FetchXml" display-name-key="FetchXml Query" description-key="FetchXml query with [RECORDID] placeholder" of-type="Multiple" usage="input" required="false" />
      <property-set name="ColumnLayoutJson" display-name-key="Column Layout" description-key="JSON column configuration" of-type="Multiple" usage="input" required="false" />
    </data-set>
    
    <!-- Properties -->
    <property name="ItemsPerPage" display-name-key="Items Per Page" description-key="Records per page (default: 50)" of-type="Whole.None" usage="input" required="false" />
    
    <property name="DebugMode" display-name-key="Debug Mode" description-key="Enable debug mode" of-type="Enum" usage="input" required="false">
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
      <uses-feature name="Utility" required="true" />
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
    
    constructor() {
        this._container = document.createElement("div");
        this._container.style.width = "100%";
        this._container.style.height = "100%";
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        
        // Check if running in test harness
        this._isTestHarness = this.isTestHarness();
        
        // Set page size from input parameter or default
        const itemsPerPage = context.parameters.ItemsPerPage?.raw;
        this._pageSize = itemsPerPage ? itemsPerPage : 50;
        
        console.log("FetchXmlDetailsList initialized", {
            isTestHarness: this._isTestHarness,
            pageSize: this._pageSize,
            hasDataset: !!(context.parameters as any).dataSet
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        
        try {
            // Check if we have dataset (view/subgrid mode)
            const dataset = (context.parameters as any).dataSet;
            
            if (dataset && dataset.records && !this._isTestHarness) {
                console.log("Using dataset mode (view/subgrid)");
                this.loadFromDataset(dataset);
            } else {
                // Re-check if test harness on each update
                this._isTestHarness = this.isTestHarness();
                
                // Check if we have valid FetchXml and ColumnLayout
                const fetchXml = context.parameters.FetchXml?.raw || "";
                const columnLayout = context.parameters.ColumnLayoutJson?.raw || "";
                
                // If in test harness OR missing required params, use sample data
                if (this._isTestHarness || !fetchXml || !columnLayout || 
                    fetchXml.trim() === "" || columnLayout.trim() === "" || 
                    columnLayout === "val") {
                    console.log("Using sample data (test harness or missing parameters)");
                    this.loadSampleData();
                } else {
                    this.loadDynamicsData();
                }
            }
        } catch (error) {
            console.error("Error updating view:", error);
            this.showError("Failed to load data: " + (error as Error).message);
        }
    }

    private isTestHarness(): boolean {
        try {
            const webAPI = (this._context as any).webAPI;
            return !webAPI || typeof webAPI.retrieveMultipleRecords !== 'function';
        } catch {
            return true;
        }
    }

    private loadFromDataset(dataset: any): void {
        console.log("Loading from dataset", dataset);
        
        const debugMode = this.getDebugMode();
        
        // Get records from dataset
        const records = dataset.sortedRecordIds.map((id: string) => {
            return dataset.records[id];
        });

        // Build columns from dataset columns
        const columns = dataset.columns.map((col: any) => ({
            key: col.name,
            fieldName: col.name,
            name: col.displayName,
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        }));

        // Get pagination info
        const hasNextPage = dataset.paging.hasNextPage;
        const hasPreviousPage = dataset.paging.hasPreviousPage;
        this._totalRecords = dataset.paging.totalResultCount || records.length;

        if (debugMode) {
            console.log("Dataset records:", records);
            console.log("Dataset columns:", columns);
            console.log("Dataset paging:", dataset.paging);
        }

        const props: any = {
            context: this._context,
            items: records,
            columns: columns,
            primaryEntityName: dataset.getTargetEntityType(),
            isDebugMode: debugMode,
            baseD365Url: this.getBaseEnvironmentUrl(),
            pagination: {
                currentPage: this._currentPage,
                pageSize: this._pageSize,
                totalRecords: this._totalRecords,
                onPageChange: (page: number) => this.handleDatasetPageChange(page, dataset)
            }
        };

        this.renderDetailsList(props);
    }

    private handleDatasetPageChange(page: number, dataset: any): void {
        if (page > this._currentPage && dataset.paging.hasNextPage) {
            dataset.paging.loadNextPage();
        } else if (page < this._currentPage && dataset.paging.hasPreviousPage) {
            dataset.paging.loadPreviousPage();
        }
        this._currentPage = page;
    }

    private loadSampleData(): void {
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
    }

    private async loadDynamicsData(): Promise<void {
        const debugMode = this.getDebugMode();
        
        if (debugMode) {
            debugger;
        }

        let fetchXml = this._context.parameters.FetchXml?.raw || "";
        const columnLayoutJson = this._context.parameters.ColumnLayoutJson?.raw || "[]";
        const recordIdPlaceholder = this._context.parameters.RecordIdPlaceholder?.raw || "[RECORDID]";
        const overriddenRecordIdFieldName = this._context.parameters.OverriddenRecordIdFieldName?.raw;

        if (!fetchXml || fetchXml.trim() === "") {
            console.warn("No FetchXml provided, loading sample data instead");
            this.loadSampleData();
            return;
        }

        let recordId = this.getRecordId(overriddenRecordIdFieldName || undefined);
        
        if (!recordId) {
            this.showError("Could not determine record ID");
            return;
        }

        fetchXml = fetchXml.replace(new RegExp(recordIdPlaceholder, 'g'), recordId);
        const paginatedFetchXml = this.addPaginationToFetchXml(fetchXml, this._currentPage, this._pageSize);

        if (debugMode) {
            console.log("DynamicDetailsList fetchXml:", paginatedFetchXml);
            console.log("DynamicDetailsList columnLayout:", columnLayoutJson);
        }

        let columns: any[];
        try {
            const cleanedJson = columnLayoutJson.trim();
            if (!cleanedJson || cleanedJson === "" || cleanedJson === "val") {
                console.warn("Invalid or empty ColumnLayoutJson, loading sample data");
                this.loadSampleData();
                return;
            }
            columns = JSON.parse(cleanedJson);
        } catch (error) {
            console.error("Failed to parse ColumnLayoutJson:", error);
            console.warn("Loading sample data instead");
            this.loadSampleData();
            return;
        }

        try {
            const entityName = this.extractEntityNameFromFetchXml(fetchXml);
            const webAPI = (this._context as any).webAPI;
            const encodedFetchXml = encodeURIComponent(paginatedFetchXml);
            const result = await webAPI.retrieveMultipleRecords(
                entityName,
                `?fetchXml=${encodedFetchXml}`
            );

            if (debugMode) {
                console.log("webAPI.retrieveMultipleRecords result:", result);
            }

            this._totalRecords = await this.getTotalRecordCount(entityName, fetchXml);

            if (debugMode) {
                console.log("Total records:", this._totalRecords);
                console.log("Current page:", this._currentPage);
                console.log("Page size:", this._pageSize);
            }

            const props: any = {
                context: this._context,
                items: result.entities || [],
                columns: columns,
                primaryEntityName: entityName,
                fetchXml: paginatedFetchXml,
                isDebugMode: debugMode,
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
            console.error("Error executing FetchXml:", error);
            this.showError("Failed to execute FetchXml: " + (error as Error).message);
        }
    }

    private getDebugMode(): boolean {
        const debugValue = this._context.parameters.DebugMode?.raw;
        return debugValue === "1";
    }

    private getRecordId(overriddenFieldName?: string): string | null {
        try {
            if (overriddenFieldName) {
                const lookupValue = (this._context as any).parameters[overriddenFieldName];
                if (lookupValue && lookupValue.raw && lookupValue.raw[0]) {
                    return lookupValue.raw[0].id;
                }
            }

            const pageContext = (this._context as any).page;
            if (pageContext && pageContext.entityId) {
                return pageContext.entityId;
            }

            const mode = this._context.mode as any;
            if (mode && mode.contextInfo && mode.contextInfo.entityId) {
                return mode.contextInfo.entityId;
            }

            return null;
        } catch (error) {
            console.error("Error getting record ID:", error);
            return null;
        }
    }

    private getBaseEnvironmentUrl(): string {
        try {
            const pageContext = (this._context as any).page;
            
            if (pageContext && typeof pageContext.getClientUrl === 'function') {
                try {
                    return pageContext.getClientUrl();
                } catch (e) {
                    console.debug("getClientUrl failed, trying alternative methods");
                }
            }
            
            if (this._context && (this._context as any).client) {
                const client = (this._context as any).client;
                if (client.getClient && typeof client.getClient === 'function') {
                    const clientInfo = client.getClient();
                    if (clientInfo && clientInfo.baseUrl) {
                        return clientInfo.baseUrl;
                    }
                }
            }
            
            if (typeof window !== 'undefined' && window.location) {
                return window.location.origin;
            }
            
            return "";
        } catch (error) {
            console.debug("Could not get base environment URL (test harness mode):", error);
            return typeof window !== 'undefined' && window.location ? window.location.origin : "";
        }
    }

    private extractEntityNameFromFetchXml(fetchXml: string): string {
        const match = fetchXml.match(/<entity\s+name=['"]([^'"]+)['"]/);
        return match ? match[1] : "entity";
    }

    private addPaginationToFetchXml(fetchXml: string, page: number, pageSize: number): string {
        let xml = fetchXml.replace(/\s+count=['"][^'"]*['"]/g, '');
        xml = xml.replace(/\s+page=['"][^'"]*['"]/g, '');
        xml = xml.replace(/\s+paging-cookie=['"][^'"]*['"]/g, '');

        xml = xml.replace(
            /<fetch/,
            `<fetch count="${pageSize}" page="${page}"`
        );

        return xml;
    }

    private async getTotalRecordCount(entityName: string, fetchXml: string): Promise<number> {
        try {
            let countFetchXml = fetchXml.replace(/<fetch[^>]*>/, '<fetch aggregate="true">');
            countFetchXml = countFetchXml.replace(
                /<entity[^>]*>/,
                (match) => match + `<attribute name="${entityName}id" alias="count" aggregate="count" />`
            );

            const webAPI = (this._context as any).webAPI;
            const encodedFetchXml = encodeURIComponent(countFetchXml);
            const result = await webAPI.retrieveMultipleRecords(
                entityName,
                `?fetchXml=${encodedFetchXml}`
            );
            
            if (result && result.entities && result.entities.length > 0) {
                return parseInt(result.entities[0].count) || 0;
            }
            
            return 0;
        } catch (error) {
            console.warn("Could not get total count:", error);
            return 0;
        }
    }

    private handlePageChange(page: number): void {
        this._currentPage = page;
        this.updateView(this._context);
    }

    private renderDetailsList(props: any): void {
        while (this._container.firstChild) {
            this._container.removeChild(this._container.firstChild);
        }

        const element = React.createElement(DynamicDetailsList, props);
        ReactDOM.render(element, this._container);
    }

    private showError(message: string): void {
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
        if (this._container) {
            ReactDOM.unmountComponentAtNode(this._container);
        }
    }
}


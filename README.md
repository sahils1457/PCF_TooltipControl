<?xml version="1.0" encoding="utf-8" ?>
<manifest>
  <control namespace="DonSchia" constructor="FetchXmlDetailsList" version="2.0.1" display-name-key="FetchXml DetailsList" description-key="FluentUI DetailsList for subgrids with FetchXml, pagination, and enhanced features" control-type="standard">
    
    <!-- Property used to store the bound field - can be any text field -->
    <property name="sampleProperty" display-name-key="Bound Field" description-key="Bind this control to any text field on the form" of-type="SingleLine.Text" usage="bound" required="true" />
    
    <!-- FetchXml Query Parameter -->
    <property name="FetchXml" display-name-key="FetchXml Query" description-key="Complete FetchXml query with [RECORDID] placeholder" of-type="Multiple" usage="input" required="false" />
    
    <!-- Column Layout JSON Parameter -->
    <property name="ColumnLayoutJson" display-name-key="Column Layout (JSON)" description-key="JSON array defining column layout and configuration" of-type="Multiple" usage="input" required="false" />
    
    <!-- Record ID Placeholder -->
    <property name="RecordIdPlaceholder" display-name-key="Record ID Placeholder" description-key="Placeholder text in FetchXml to replace with record ID (default: [RECORDID])" of-type="SingleLine.Text" usage="input" required="false" />
    
    <!-- Override Record ID Field Name -->
    <property name="OverriddenRecordIdFieldName" display-name-key="Override Record ID Field" description-key="Optional: Use a lookup field on the form instead of current record ID" of-type="SingleLine.Text" usage="input" required="false" />
    
    <!-- Items Per Page for Pagination -->
    <property name="ItemsPerPage" display-name-key="Items Per Page" description-key="Number of items to display per page (default: 50)" of-type="Whole.None" usage="input" required="false" />
    
    <!-- Debug Mode -->
    <property name="DebugMode" display-name-key="Debug Mode" description-key="Enable debug mode (On/Off)" of-type="Enum" usage="input" required="false">
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
    
    constructor() {
        // Constructor should be empty or minimal
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
            
            // Set container styles
            if (this._container) {
                this._container.style.width = "100%";
                this._container.style.height = "100%";
            }
            
            // Check if running in test harness
            this._isTestHarness = this.isTestHarness();
            
            // Access parameters safely
            const params = context.parameters as any;
            
            // Set page size from input parameter or default
            const itemsPerPage = params.ItemsPerPage?.raw;
            this._pageSize = itemsPerPage ? itemsPerPage : 50;
            
            console.log("FetchXmlDetailsList initialized", {
                isTestHarness: this._isTestHarness,
                pageSize: this._pageSize
            });
        } catch (error) {
            console.error("Error in init:", error);
            this.showError("Initialization failed: " + (error as Error).message);
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        try {
            this._context = context;
            
            // Re-check if test harness on each update
            this._isTestHarness = this.isTestHarness();
            
            // Access parameters safely
            const params = context.parameters as any;
            
            // Check if we have valid FetchXml and ColumnLayout
            const fetchXml = params.FetchXml?.raw || "";
            const columnLayout = params.ColumnLayoutJson?.raw || "";
            
            // If in test harness OR missing required params, use sample data
            if (this._isTestHarness || !fetchXml || !columnLayout || 
                fetchXml.trim() === "" || columnLayout.trim() === "" || 
                columnLayout === "val") {
                console.log("Using sample data (test harness or missing parameters)");
                this.loadSampleData();
            } else {
                this.loadDynamicsData();
            }
        } catch (error) {
            console.error("Error updating view:", error);
            this.showError("Failed to load data: " + (error as Error).message);
        }
    }

    private isTestHarness(): boolean {
        // Check if we're in the test harness by checking for WebApi
        try {
            const webAPI = (this._context as any)?.webAPI;
            return !webAPI || typeof webAPI.retrieveMultipleRecords !== 'function';
        } catch {
            return true;
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

    private async loadDynamicsData(): Promise<void> {
        try {
            const debugMode = this.getDebugMode();
            
            if (debugMode) {
                debugger;
            }

            // Access parameters safely
            const params = this._context.parameters as any;
            
            // Get parameters
            let fetchXml = params.FetchXml?.raw || "";
            const columnLayoutJson = params.ColumnLayoutJson?.raw || "[]";
            const recordIdPlaceholder = params.RecordIdPlaceholder?.raw || "[RECORDID]";
            const overriddenRecordIdFieldName = params.OverriddenRecordIdFieldName?.raw;

            if (!fetchXml || fetchXml.trim() === "") {
                console.warn("No FetchXml provided, loading sample data instead");
                this.loadSampleData();
                return;
            }

            // Get record ID
            let recordId = this.getRecordId(overriddenRecordIdFieldName || undefined);
            
            if (!recordId) {
                this.showError("Could not determine record ID");
                return;
            }

            // Replace placeholder with actual record ID
            fetchXml = fetchXml.replace(new RegExp(recordIdPlaceholder, 'g'), recordId);

            // Add pagination to FetchXml
            const paginatedFetchXml = this.addPaginationToFetchXml(fetchXml, this._currentPage, this._pageSize);

            if (debugMode) {
                console.log("DynamicDetailsList fetchXml:", paginatedFetchXml);
                console.log("DynamicDetailsList columnLayout:", columnLayoutJson);
            }

            // Parse columns
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

            // Execute FetchXml
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

            // Get total record count for pagination
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
        try {
            const params = this._context?.parameters as any;
            const debugValue = params?.DebugMode?.raw;
            return debugValue === "1";
        } catch {
            return false;
        }
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
            const pageContext = (this._context as any)?.page;
            
            if (pageContext && typeof pageContext.getClientUrl === 'function') {
                try {
                    return pageContext.getClientUrl();
                } catch (e) {
                    console.debug("getClientUrl failed");
                }
            }
            
            if (this._context && (this._context as any).client) {
                const client = (this._context as any).client;
                if (client?.getClient && typeof client.getClient === 'function') {
                    const clientInfo = client.getClient();
                    if (clientInfo?.baseUrl) {
                        return clientInfo.baseUrl;
                    }
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
            
            if (result?.entities && result.entities.length > 0) {
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
        try {
            if (!this._container) {
                console.error("Container not initialized");
                return;
            }

            // Clear container
            while (this._container.firstChild) {
                this._container.removeChild(this._container.firstChild);
            }

            // Create React element and render
            const element = React.createElement(DynamicDetailsList, props);
            ReactDOM.render(element, this._container);
        } catch (error) {
            console.error("Error rendering details list:", error);
            this.showError("Failed to render: " + (error as Error).message);
        }
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

/// <reference path="../Typing/ControlFramework.d.ts"/>

import {IInputs, IOutputs} from "./generated/ManifestTypes";

/*
import { debug } from "util";
import { array } from "prop-types";

//https://github.com/redbooth/free-file-icons
 * The MIT License (MIT)
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy,
modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the
Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */


class Downloadfile implements ComponentFramework.FileObject{
    fileContent: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
}

class EntityReference {
    id: string;
    typeName: string;
    constructor(typeName: string, id: string) {
        this.id = id;
        this.typeName = typeName;
    }
}


class Item { 
    attachmentId: EntityReference;
    groupId: EntityReference;
    groupName: string;
    name: string;
    size: number;
    extension: string;
    entityType: string;
    constructor(groupId: EntityReference,attachmentId: EntityReference, groupName: string,  name: string, extension: string, size: number) {
        this.attachmentId = attachmentId;
        this.groupId = groupId;
        this.groupName = groupName;
        this.name = name;
        this.size = size;
        this.extension = extension;
    }
}



export class FileExplorer implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _groupElements: HTMLElement[] = [];
    private _itemElements: HTMLElement[] = [];

    
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

    /**
     * Render table with items
     * @param items
     * items to render
     */
    private renderTable(items: Item[]) {

        if (items.length > 0) {
            let divControl: HTMLDivElement = document.createElement("div");
            divControl.className = "control";
            this._container.appendChild(divControl);
            let groupId = '-';
            let i: number = 0;
            do {                
                if (items[i].groupId != null && items[i].groupId.id != groupId) {
                    //add new group
                    this.addGroup(divControl, items[i]);
                    groupId = items[i].groupId.id;
                }
                //add item
                this.addItem(divControl, items[i]);
                i++;                
            } while (i < items.length);
        }
    }

    /**
     * Add group
     * @param control
     * Control where group will be added
     * @param item
     * Item to render
     */
    private addGroup(control: HTMLDivElement, item: Item) {
        //create title
        let divTitle: HTMLDivElement = <HTMLTableElement>document.createElement("div");
        divTitle.className = "group";
        if (item.groupName == '' ) {
            divTitle.innerText = "Attachments";
        }
        else
        {
            divTitle.innerText = item.groupName;
            divTitle.onclick =
                (e => {
                    this.openEntity(item.groupId);
                });
            this._groupElements.push(divTitle);
        }
        control.appendChild(divTitle);
    }

    
    /**
     * Add item
     * @param control
     * Control where item will be added
     * @param item
     * Item to render
     */
    private addItem(control: HTMLDivElement, item: Item) {

        let divItem: HTMLDivElement = document.createElement("div");
        divItem.className = "item";
        divItem.onclick = (e => {
                this.onClickAttachment(item.attachmentId);
        }

        );
        control.appendChild(divItem);
        this._itemElements.push(divItem);

        let tblItem: HTMLTableElement = document.createElement("table");
        divItem.appendChild(tblItem);

        let imgRow: HTMLTableRowElement = tblItem.insertRow();
        let imgCell: HTMLTableCellElement = imgRow.insertCell();
        let img: HTMLImageElement = <HTMLImageElement>document.createElement("img");

        //add image
        let res = this._context.resources.getResource(item.extension + ".png",
            this.setImage.bind(this, img, "png"),
            this.showError.bind(this)
            );
        imgCell.appendChild(img);

        //add subtitle
        let subTitleRow: HTMLTableRowElement = tblItem.insertRow();
        let subTitleCell: HTMLTableCellElement = subTitleRow.insertCell();
        subTitleCell.innerText = item.name;        
    }

    /**
     * When clicked on attachment, start download
     * @param reference
     * EntityReference to download
     */
    private onClickAttachment(reference: EntityReference): void {       
        this.downloadAttachment(reference).then(f => {
            this._context.navigation.openFile(f);   }
        );        
    }

    /**
     * Open entity form
     * @param reference
     * EntityReference to open
     */
    private openEntity(reference: EntityReference) :void { 
        let entityFormOptions = <any>{};
        entityFormOptions["entityName"] = reference.typeName
        entityFormOptions["entityId"] = reference.id;
        (<any>this._context).navigation.openForm(entityFormOptions);
    }

    /**
     * Download the attachment
     * @param reference
     * EntityReference to download
     */
    private downloadAttachment(reference: EntityReference): Promise<Downloadfile> {
        debugger;
        return this._context.webAPI.retrieveRecord(reference.typeName, reference.id).then(
            function success(result) {
                let file: Downloadfile = new Downloadfile();
                file.fileContent =
                    reference.typeName == "annotation" ? result["documentbody"] : result["body"];
                file.fileName = result["filename"];
                file.fileSize = result["filesize"];
                file.mimeType = result["mimetype"];
                return file;               
            });
    }

    /**
     * Add image url
     * @param element
     * Element that contains the url
     * @param fileType
     * Type of the file
     * @param fileContent
     * Content of the file
     */
    private setImage(element: HTMLImageElement, fileType: string, fileContent: string): void {
        let imageUrl: string = this.generateImageSrcUrl(fileType, fileContent);
        element.src = imageUrl;     
    }
    /***
     * create source of file
     * @param fileType
     * filetype
     * @param fileContent
     * content
     */
    private generateImageSrcUrl(fileType: string, fileContent: string): string {
        return "data:image/" + fileType + ";base64, " + fileContent;
    }

    /**
     * log error
     */
    private showError(): void {
        console.log('error while downloading .png');
    }

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
    {       
		// Add control initialization code
        // assigning environment variables. 
        this._context = context;
        this._container = container;

        //when we associate to lookup then lookup record's attachments and e-mail attachments
        if (this._context.parameters.hidden.type.indexOf('Lookup') == 0) {
            let ref = context.parameters.hidden.raw;
            if (ref && ref.length > 0) {
                let reference: EntityReference = new EntityReference(
                    ref[0].LogicalName,
                    ref[0].id
                );
                this.getAttachments(reference).then(r => this.renderTable(r));
                this.getRelatedEmailAttachments(reference).then(r => this.renderTable(r));
            }
        }
        else {
            //get current record's attachments and e-mail attachments
            let reference: EntityReference = new EntityReference(
                (<any>context).page.entityTypeName,
                (<any>context).page.entityId
            )
            if ((<any>context).page.entityId != null) {
                this.getAttachments(reference).then(r => this.renderTable(r));
                this.getRelatedEmailAttachments(reference).then(r => this.renderTable(r));
            }
        }
	}

    /**
     * Get the e-mail attachmentd
     * @param reference
     * Object that contains e-mails
     */
    private getRelatedEmailAttachments(reference : EntityReference): Promise<Item[]> {
        let xml = '<fetch><entity name="activitymimeattachment"><attribute name="activityid"/><attribute name="activitymimeattachmentid"/><attribute name="activitysubject"/><attribute name="filename"/><attribute name="filesize"/><link-entity name="email" from="activityid" to="objectid" link-type="inner"><attribute name="createdon"/><attribute name="directioncode"/><attribute name="activityid"/><attribute name="regardingobjectid"/><filter type="and"><condition attribute="regardingobjectid" operator="eq" value="[OBJECTID]"/></filter></link-entity></entity></fetch >';
        xml = xml.replace('[OBJECTID]', reference.id);
        let query = '?fetchXml=' + encodeURIComponent(xml);        
        return this._context.webAPI.retrieveMultipleRecords("activitymimeattachment", query).then(
            function success(result) {
                let items: Item[] = [];
                for (let i = 0; i < result.entities.length; i++) {
                    let ent = result.entities[i];
                    let it = new Item(
                        new EntityReference("email", ent["email1.activityid"].toString()),
                        new EntityReference("activitymimeattachment", ent["activitymimeattachmentid"].toString()),
                        new Date(ent["email1.createdon"]).toLocaleDateString() +
                        " " +
                        ent["activitysubject"]
                        , ent["filename"].split('.')[0],
                        ent["filename"].split('.')[1],
                        ent["filesize"]);
                    items.push(it);
                }
                return items;
            }
            , function (error) {
                console.log(error.message);
                let items: Item[] = [];
                return items;
            }
        );

    }

    /**
     * Get attachment items
     * @param reference
     * EntityReference to get the attachments of
     */
    private getAttachments(reference: EntityReference): Promise<Item[]> {
        let query = "?$select=annotationid,filename,filesize,mimetype&$filter=filename ne null and _objectid_value eq " + reference.id + " and objecttypecode eq '" + reference.typeName + "' &$orderby=createdon desc";       
        return this._context.webAPI.retrieveMultipleRecords("annotation", query).then(
            function success(result) {
                                    let items: Item[] = [];
                                    for (let i = 0; i < result.entities.length; i++) {
                                        let ent = result.entities[i];
                                        let it = new Item(
                                            null,
                                            new EntityReference("annotation", ent["annotationid"].toString()),
                                            "",
                                            ent["filename"].split('.')[0],
                                            ent["filename"].split('.')[1],                                            
                                            ent["filesize"]);
                                        items.push(it);
                                    }
                                    return items;
                                    }
           , function (error) {
               console.log(error.message);
               let items: Item[] = [];
               return items;
           }
        );

    }


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
    }

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
    {
        if (this._itemElements != null) {
            for (let i = 0; i < this._itemElements.length; i++) {
                this._itemElements[i].removeEventListener("click",
                    (e => {
                        this.onClickAttachment
                    }));
            }
        }
        if (this._groupElements != null) {
            for (let i = 0; i < this._groupElements.length; i++) {
                this._groupElements[i].removeEventListener("click",
                    (e => {
                        this.openEntity
                    }) );
            }
        }
    }
    
}
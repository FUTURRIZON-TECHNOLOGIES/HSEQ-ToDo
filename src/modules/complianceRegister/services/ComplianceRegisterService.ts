import { IComplianceItem, ILookupOption, IAttachment, ILookupSets } from '../models/IComplianceItem';

export class ComplianceService {
  private siteUrl: string;
  private _lookups: ILookupSets & { users?: ILookupOption[] } | undefined;

  constructor(siteUrl: string) {
    this.siteUrl = siteUrl.replace(/\/$/, '');
  }

  private async getRequestDigest(): Promise<string> {
    const response = await fetch(`${this.siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { Accept: 'application/json;odata=verbose' }
    });
    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }

  // ─── List Items ──────────────────────────────────────────────────────────────

  public async getItems(
    top: number = 15, 
    skip: number = 0, 
    filters: Record<string, string> = {},
    sortColumn: string = 'Id',
    isAscending: boolean = false
  ): Promise<{ items: IComplianceItem[]; totalCount: number }> {

    // Build filter string
    const filterParts: string[] = [];
    for (const key in filters) {
      if (Object.prototype.hasOwnProperty.call(filters, key)) {
        const value: string = filters[key];
        if (value && value.trim()) {
          filterParts.push(`substringof('${value.replace(/'/g, "''")}', ${key})`);
        }
      }
    }
    const filterStr = filterParts.length > 0 ? `&$filter=${filterParts.join(' and ')}` : '';

    const countUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/ItemCount`;
    const orderStr = `${sortColumn} ${isAscending ? 'asc' : 'desc'}`;
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items?$top=${top}&$skip=${skip}&$orderby=${orderStr}${filterStr}`;

    const [itemsResp, countResp] = await Promise.all([
      fetch(url, { headers: { Accept: 'application/json;odata=verbose' } }),
      fetch(countUrl, { headers: { Accept: 'application/json;odata=verbose' } })
    ]);

    if (!itemsResp.ok) {
        const errorText = await itemsResp.text();
        throw new Error(`SharePoint API Error: ${errorText}`);
    }

    const itemsData = await itemsResp.json();
    const countData = await countResp.json();

    const rawItems = itemsData.d ? itemsData.d.results : (itemsData.value || []);
    
    // For debugging: dump the first item's keys into an error if length > 0
    // if (rawItems.length > 0) {
    //   throw new Error("Available columns: " + Object.keys(rawItems[0]).join(', '));
    // }

    const items: IComplianceItem[] = rawItems.map((raw: any) => this.mapRawToItem(raw));
    return { items, totalCount: countData.d ? countData.d.ItemCount : (countData.value || 0) };
  }

  public async getItem(id: number): Promise<IComplianceItem> {
    // Use odata=verbose + data.d parsing — EXACTLY matching getItems which works reliably.
    // With nometadata the Id or key field names can differ; verbose is the safe consistent format.
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${id})`;
    const resp = await fetch(url, { headers: { Accept: 'application/json;odata=verbose' } });
    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(`SharePoint getItem error ${resp.status}: ${err}`);
    }
    const data = await resp.json();
    // Single item: wrapped as data.d (not data.d.results like collections)
    const raw = (data.d) ? data.d : data;
    return this.mapRawToItem(raw);
  }


  /** Fetches only the Author (Created By) display name for a single item. Used by the ZIP download. */
  public async getItemAuthor(id: number): Promise<string> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${id})?$select=Author/Title&$expand=Author`;
      const resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!resp.ok) return '';
      const data = await resp.json();
      return data?.Author?.Title || '';
    } catch {
      return '';
    }
  }

  private mapRawToItem(raw: any): IComplianceItem {
    // Helper to find key ignoring case and _x0020_ (SharePoint strictness)
    const keys = Object.keys(raw);
    const getVal = (names: string[]) => {
      for (const n of names) {
        let exactMatch = null;
        for (const k of keys) {
          const cleanK = k.toLowerCase().replace(/_x0020_/g, '');
          if (cleanK === n.toLowerCase()) return raw[k];
          if (cleanK === n.toLowerCase() + 'id') exactMatch = raw[k];
        }
        if (exactMatch !== null) return exactMatch;
      }
      return null;
    };

    const expDateStr = getVal(['ExpiryDate', 'Expiry_x0020_Date']) || null;
    let daysRemaining: string | number = '';
    let status = '';
    let expired = 'No';

    if (expDateStr) {
      const expD = new Date(expDateStr);
      const today = new Date();
      const diffTime = expD.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      daysRemaining = diffDays;
      if (diffDays < 0) {
        status = 'Expired';
        expired = 'Yes';
      } else {
        status = 'Active';
      }
    }

    const typeId = getVal(['ComplianceTypeId', 'ComplianceType']) || null;
    const finalTypeId = (typeId && typeof typeId === 'object' && typeId.Id) ? typeId.Id : typeId;

    let personName = getVal(['PersonName', 'Name', 'Person']) || '';
    let genericNameStr = raw.Name || raw.Company || personName || '';
    
    // Reverse Engineering Lookup Strings into Strict IDs
    let projectId = raw.ProjectId || null, empId = raw.EmployeeId || null, subId = raw.SubcontractorId || null, workId = raw.WorkerId || null, busId = raw.BusinessId || null;
    
    if (this._lookups) {
        let rawProj = getVal(['Project']) || genericNameStr;
        let foundProj = this._lookups.projects.find((x:any) => x.text === rawProj);
        if (foundProj) projectId = foundProj.key;

        let rawEmp = getVal(['Employee']) || genericNameStr;
        let foundEmp = this._lookups.employees.find((x:any) => x.text === rawEmp);
        if (foundEmp) empId = foundEmp.key;
        
        let rawSub = getVal(['Subcontractor']) || genericNameStr;
        let foundSub = this._lookups.subcontractors.find((x:any) => x.text === rawSub);
        if (foundSub) subId = foundSub.key;
        
        let rawWork = getVal(['Worker']) || genericNameStr;
        let foundWork = this._lookups.workers.find((x:any) => x.text === rawWork);
        if (foundWork) workId = foundWork.key;
        
        let rawBus = getVal(['Business']) || genericNameStr;
        let foundBus = this._lookups.businesses.find((x:any) => x.text === rawBus);
        if (foundBus) busId = foundBus.key;
    }

    // Extract Position from concatenated Employee/Worker strings (part[1])
    let position = '';
    const complianceFor = getVal(['ComplianceFor']) || '';
    if (complianceFor === 'Employee') {
      const empParts = (raw.Employee?.Employee_x0020_Name || '').split(' | ');
      position = empParts[1]?.trim() || '';
    } else if (complianceFor === 'Worker') {
      const workerParts = (raw.Worker?.Worker || '').split(' | ');
      position = workerParts[1]?.trim() || '';
    }

    // Resolve Business Profile name
    const businessProfileName = raw.Business?.Business_x0020_Profile || '';

    // Resolve the CreatedBy (Author) field — handle both nometadata (inline) and verbose (deferred) responses
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const authorObj: any = raw.Author;
    const authorTitle = (authorObj && !authorObj.__deferred)
      ? (authorObj.Title || authorObj.LoginName || '')
      : '';

    return {
      Id: raw.Id || raw.ID,
      ComplianceFor: complianceFor as any,
      ComplianceTypeId: finalTypeId,
      ComplianceTypeName: raw.ComplianceType?.Title || '',
      MainBusinessProfileId: raw.BusinessProfileId || raw.Business_x0020_ProfileId || null,
      BusinessId: busId,
      BusinessName: raw.Business?.Business_x0020_Profile || raw.Business || genericNameStr,
      EmployeeId: empId,
      EmployeeName: raw.Employee?.Employee_x0020_Name || raw.Employee || genericNameStr,
      ProjectId: projectId,
      ProjectTitle: raw.Project?.Project_x0020_Title || raw.Project || genericNameStr,
      SubcontractorId: subId,
      SubcontractorName: raw.Subcontractor?.Company_x0020_Name || raw.Subcontractor || genericNameStr,
      WorkerId: workId,
      WorkerName: raw.Worker?.Worker || raw.Worker || genericNameStr,
      IsBooking: getVal(['IsBooking']) || false,
      BookingDate: getVal(['BookingDate', 'DateTime', 'Date']) || null,
      BookedWith: getVal(['BookedWith', 'Comment']) || '',
      DocumentNumber: getVal(['DocumentNumber', 'Document_x0020_x0023_']) || '',
      IssuingAuthority: getVal(['IssuingAuthority', 'Issuing_x0020_Authority']) || '',
      IssueDate: getVal(['IssueDate', 'Issue_x0020_Date']) || null,
      RenewalNotRequired: getVal(['RenewalNotRequired', 'Renewal_x0020_Not_x0020_Required', 'RenewalRequired']) === true || getVal(['RenewalNotRequired', 'Renewal_x0020_Not_x0020_Required', 'RenewalRequired']) === 'Yes',
      ExpiryDate: expDateStr,
      Notes: raw.Notes || '',
      HasAttachments: raw.HasAttachments || false,

      EntityActive: 'Yes',
      Status: status,
      Expired: expired,
      DaysRemaining: daysRemaining === '' ? '' : daysRemaining.toString(),

      // Audit/extra fields
      CreatedBy: authorTitle,
      DateEntered: raw.Created || null,
      Position: position,
      BusinessProfile: businessProfileName
    };
  }

  private _listFieldsMapping: { [key: string]: string } = {};

  private async getInternalFieldMap(): Promise<{ [key: string]: string }> {
      try {
          if (Object.keys(this._listFieldsMapping).length > 0) return this._listFieldsMapping;
          const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/fields?$select=Title,InternalName`;
          const resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
          const data = await resp.json();
          const map: { [key: string]: string } = {};
          (data.value || []).forEach((f: any) => {
              if (f.Title) {
                 map[f.Title] = f.InternalName;
                 map[f.Title.toLowerCase()] = f.InternalName;
              }
          });
          this._listFieldsMapping = map;
          return map;
      } catch (e) {
          return {};
      }
  }

  private buildPayload(item: Partial<IComplianceItem>, fMap: { [key:string]: string }): any {
    const payload: any = {
      '__metadata': { type: 'SP.Data.ComplianceRegisterListItem' }
    };
    
    const setF = (title: string, fallback: string, value: any) => {
       if (value === undefined) return;
       const iName = fMap[title] || fMap[title.toLowerCase()] || fallback;
       payload[iName] = value;
    };

    setF('Compliance For', 'ComplianceFor', item.ComplianceFor);
    
    if (item.ComplianceTypeId !== undefined) {
        let iName = fMap['Compliance Type'] || fMap['compliance type'] || 'ComplianceTypeId';
        if (!iName.endsWith('Id') && !iName.endsWith('ID')) iName += 'Id';
        payload[iName] = item.ComplianceTypeId;
    }

    if (item.MainBusinessProfileId !== undefined) {
        let iName = fMap['Business Profile'] || fMap['business profile'] || 'BusinessProfileId';
        if (!iName.endsWith('Id') && !iName.endsWith('ID')) iName += 'Id';
        payload[iName] = item.MainBusinessProfileId ?? null;
    }

    setF('Is Booking?', 'IsBooking', item.IsBooking);
    
    // Renewal Check
    if (item.RenewalNotRequired !== undefined) {
        if (fMap['Renewal Required']) {
            payload[fMap['Renewal Required']] = !item.RenewalNotRequired;
        } else if (fMap['Renewal Not required'] || fMap['renewal not required']) {
            payload[fMap['Renewal Not required'] || fMap['renewal not required']] = item.RenewalNotRequired;
        } else {
            payload['RenewalRequired'] = !item.RenewalNotRequired;
        }
    }

    setF('Document #', 'DocumentNumber', item.DocumentNumber);
    setF('Issuing Authority', 'IssuingAuthority', item.IssuingAuthority);
    setF('Issue Date', 'IssueDate', item.IssueDate);
    setF('Expiry Date', 'ExpiryDate', item.ExpiryDate);
    setF('Notes', 'Notes', item.Notes);
    setF('Booked With', 'BookedWith', item.BookedWith);
    setF('Booking Date', 'BookingDate', item.BookingDate);

    // Company is ALWAYS the Business Profile value
    const businessProfileText = this._lookups?.businesses.find((p:any) => p.key === item.MainBusinessProfileId)?.text || '';


    let nameText = '';      // goes to Name column (just the entity's own name)
    let fullEntityText = ''; // goes to the specific entity column (full concatenated text)

    if (item.ComplianceFor === 'Business') {
      // For Business: Name = Business Profile, Company = Business Profile
      nameText = businessProfileText;
      fullEntityText = businessProfileText;
    } else if (item.ComplianceFor === 'Employee') {
      fullEntityText = this._lookups?.employees.find((p:any) => p.key === item.EmployeeId)?.text || '';
      // Name = just the employee name (first part before ' | ')
      nameText = fullEntityText.split(' | ')[0].trim();
    } else if (item.ComplianceFor === 'Project') {
      fullEntityText = this._lookups?.projects.find((p:any) => p.key === item.ProjectId)?.text || '';
      // Name = just the project title (first part before ' | ')
      nameText = fullEntityText.split(' | ')[0].trim();
    } else if (item.ComplianceFor === 'Subcontractor') {
      fullEntityText = this._lookups?.subcontractors.find((p:any) => p.key === item.SubcontractorId)?.text || '';
      // Name = just the company name (first part before ' | ')
      nameText = fullEntityText.split(' | ')[0].trim();
    } else if (item.ComplianceFor === 'Worker') {
      fullEntityText = this._lookups?.workers.find((p:any) => p.key === item.WorkerId)?.text || '';
      // Name = just the worker name (first part before ' | ')
      nameText = fullEntityText.split(' | ')[0].trim();
    }

    if (item.ComplianceFor && nameText) {
       // Worker: Company = employer company (part[2] of concatenated text)
       // Subcontractor: Company = company name (part[0]) same as Name
       let companyText = businessProfileText; // default = Business Profile
       if (item.ComplianceFor === 'Worker') {
         const workerParts = fullEntityText.split(' | ');
         companyText = workerParts[2]?.trim() || businessProfileText;
       } else if (item.ComplianceFor === 'Subcontractor') {
         companyText = nameText; // Company = same as Name (part[0])
       }

       payload.Title      = nameText;     // built-in Title field
       payload['Name']    = nameText;     // Name column = entity's own name
       payload['Company'] = companyText;  // Company column

       // Store the full concatenated text in the specific entity column
       if (item.ComplianceFor === 'Business')      setF('Business',      'Business',      fullEntityText);
       if (item.ComplianceFor === 'Employee')       setF('Employee',       'Employee',       fullEntityText);
       if (item.ComplianceFor === 'Project')        setF('Project',        'Project',        fullEntityText);
       if (item.ComplianceFor === 'Subcontractor')  setF('Subcontractor',  'Subcontractor',  fullEntityText);
       if (item.ComplianceFor === 'Worker')         setF('Worker',         'Worker',         fullEntityText);
    }

    return payload;
  }

  public async createItem(item: Partial<IComplianceItem>): Promise<number> {
    const fMap = await this.getInternalFieldMap();
    const digest = await this.getRequestDigest();
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items`;
    const payload = this.buildPayload(item, fMap);

    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(payload)
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Create failed: ${errText}`);
    }

    const data = await resp.json();
    return data.d.ID;
  }

  public async updateItem(id: number, item: Partial<IComplianceItem>): Promise<void> {
    const fMap = await this.getInternalFieldMap();
    const digest = await this.getRequestDigest();
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${id})`;
    const payload = this.buildPayload(item, fMap);

    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify(payload)
    });

    if (!resp.ok && resp.status !== 204) {
      const errText = await resp.text();
      throw new Error(`Update failed: ${errText}`);
    }
  }

  public async deleteItem(id: number): Promise<void> {
    const digest = await this.getRequestDigest();
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${id})`;
    await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });
  }

  // ─── Attachments ──────────────────────────────────────────────────────────────

  public async getAttachments(itemId: number): Promise<IAttachment[]> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${itemId})/AttachmentFiles`;
    const resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
    const data = await resp.json();
    return (data.value || []).map((a: any) => ({ FileName: a.FileName, ServerRelativeUrl: a.ServerRelativeUrl }));
  }

  public async uploadAttachment(itemId: number, file: File): Promise<void> {
    const digest = await this.getRequestDigest();
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
    const buffer = await file.arrayBuffer();
    await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'Content-Type': 'application/octet-stream'
      },
      body: buffer
    });
  }

  public async deleteAttachment(itemId: number, fileName: string): Promise<void> {
    const digest = await this.getRequestDigest();
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('Compliance Register')/items(${itemId})/AttachmentFiles/getByFileName('${encodeURIComponent(fileName)}')`;
    await fetch(url, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'X-RequestDigest': digest,
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });
  }

  public async getAttachmentBlob(serverRelativeUrl: string): Promise<Blob> {
    const url = `${this.siteUrl}/_api/web/getFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/$value`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Fetch attachment failed: ${resp.statusText}`);
    return await resp.blob();
  }

  // ─── Lookup Data ──────────────────────────────────────────────────────────────

  public async getAllLookups(): Promise<ILookupSets & { users: ILookupOption[] }> {
    const [compTypes, businesses, employees, projects, subcontractors, users] = await Promise.all([
      this.fetchLookup("Compliance Type", "Title"),
      this.fetchLookup("Business Profiles", "Business_x0020_Profile"),
      this.fetchEmployees(),
      this.fetchProjects(),
      this.fetchSubcontractors(),
      this.fetchUsers()
    ]);

    const workers = await this.fetchWorkers(subcontractors);

    const result = {
      complianceTypes: compTypes,
      businesses: businesses,
      employees: employees,
      projects: projects,
      subcontractors: subcontractors,
      workers: workers,
      users: users
    };
    
    this._lookups = result;
    return result;
  }

  private async fetchUsers(): Promise<ILookupOption[]> {
    try {
      const url = `${this.siteUrl}/_api/web/lists/getbytitle('User Information List')/items?$top=5000`;
      const resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!resp.ok) return [];
      const data = await resp.json();
      return (data.value || []).map((u: any) => ({ key: u.Id, text: u.Title || u.Name }));
    } catch {
      return [];
    }
  }

  private async fetchProjects(): Promise<ILookupOption[]> {
    try {
      const pUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Project')/items?$top=5000&$expand=FieldValuesAsText`;
      const pResp = await fetch(pUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!pResp.ok) return [];
      const pData = await pResp.json();

      let cUrl = `${this.siteUrl}/_api/web/lists/getbytitle('ContractType1')/items?$top=5000&$expand=FieldValuesAsText`;
      let cResp = await fetch(cUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!cResp.ok) {
         cUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Contract Type')/items?$top=5000&$expand=FieldValuesAsText`;
         cResp = await fetch(cUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      }

      const cData = cResp.ok ? await cResp.json() : { value: [] };
      const contractTypes = (cData.value || []).map((c: any) => {
        const tVals = c.FieldValuesAsText || {};
        let text = tVals.Name || tVals.Title || c.Name || c.Title;
        
        if (!text) {
           const possibleKey = Object.keys(tVals).find(k => k.toLowerCase().includes('name') || k.toLowerCase().includes('title') || k.toLowerCase().includes('contract'));
           text = possibleKey ? tVals[possibleKey] : `ColMissing:${Object.keys(tVals).slice(0,2).join(',')}`;
        }
        return { key: String(c.Id), text: text || `TypeID:${c.Id}` };
      });

      return (pData.value || []).map((p: any) => {
        const tVals = p.FieldValuesAsText || {};
        
        let pNum = String(tVals.Number || p.Number || '').trim();
        if (!pNum) {
           const numKey = Object.keys(tVals).find(k => k.toLowerCase().includes('num'));
           if (numKey) pNum = String(tVals[numKey]).trim();
        }

        const pTitle = String(tVals.Project_x0020_Title || tVals.ProjectTitle || tVals.Title || p.Project_x0020_Title || p.Title || '').trim();
        
        let rawContract = String(p.ContractType1Id || p.ContractTypeId || p.Contract_x0020_TypeId || tVals.ContractType1 || tVals.Contract_x0020_Type || tVals.ContractType || '').trim();
        if (!rawContract || rawContract === 'undefined') {
           const ctKey = Object.keys(p).find(k => k.toLowerCase().includes('contract') && k.toLowerCase().endsWith('id'));
           if (ctKey) rawContract = String(p[ctKey]).trim();
        }

        let finalContract = rawContract;
        const matched = contractTypes.find((x: any) => x.key === rawContract);
        
        if (contractTypes.length === 0) {
           finalContract = `List 404: ${rawContract}`;
        } else if (matched) {
           finalContract = matched.text;
        }
        
        if (!finalContract || finalContract === 'undefined' || finalContract === 'null' || finalContract.trim() === '') {
           const typeKeys = Object.keys(p).filter(k => k.toLowerCase().includes('type'));
           const typeVals = typeKeys.map(k => `${k}=${p[k]}`).join(',');
           finalContract = `(MISSING TYPE: ${typeVals})`;
        }

        const mainText = [pNum, pTitle].filter(Boolean).join(' - ');
        // We do NOT filter Boolean here, so if finalContract is missing, it still forces the string to render!
        const strParts = [mainText, finalContract, 'Supply Workforce'];

        return { key: p.Id || p.ID, text: strParts.join(' | ') };
      });
    } catch (e: any) {
      return [{ key: -99, text: `ERR: ${(e as Error).message}` }];
    }
  }

  private async fetchSubcontractors(): Promise<ILookupOption[]> {
    try {
      const sUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Subcontractors')/items?$top=5000&$expand=FieldValuesAsText`;
      const sResp = await fetch(sUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!sResp.ok) return [];
      const sData = await sResp.json();

      return (sData.value || []).map((s: any) => {
        const tVals = s.FieldValuesAsText || {};
        
        let comp = String(tVals.Company_x0020_Name || tVals.CompanyName || tVals.Title || s.Title || '').trim();
        if (!comp) {
            const cKey = Object.keys(tVals).find(k => k.toLowerCase().includes('company'));
            if (cKey) comp = String(tVals[cKey]).trim();
        }

        let accNum = String(tVals.Account_x0020_Number || tVals.AccountNumber || s.AccountNumber || '').trim();
        if (!accNum) {
            const accKey = Object.keys(tVals).find(k => k.toLowerCase().includes('account'));
            if (accKey) accNum = String(tVals[accKey]).trim();
        }

        let abn = String(tVals.ABN || s.ABN || '').trim();
        if (!abn) {
            const abnKey = Object.keys(tVals).find(k => k.toLowerCase().includes('abn'));
            if (abnKey) abn = String(tVals[abnKey]).trim();
        }

        let phone = String(tVals.Main_x0020_Phone || tVals.MainPhone || tVals.Phone || s.MainPhone || s.Phone || '').trim();
        if (!phone) {
            const phoneKey = Object.keys(tVals).find(k => k.toLowerCase().includes('phone') || k.toLowerCase().includes('mobile'));
            if (phoneKey) phone = String(tVals[phoneKey]).trim();
        }

        const strParts = [comp, accNum, abn, phone, 'Supply Workforce']
            .map(x => String(x).trim())
            .filter(Boolean)
            .filter(x => x !== 'undefined' && x !== 'null');
            
        return { key: s.Id || s.ID, text: strParts.join(' | ') };
      });
    } catch (e: any) {
      return [{ key: -99, text: `ERR: ${(e as Error).message}` }];
    }
  }

  private async fetchEmployees(): Promise<ILookupOption[]> {
    try {
      // 1. Fetch Position Types
      let pUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Position Types')/items?$top=5000&$expand=FieldValuesAsText`;
      let pResp = await fetch(pUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      const pData = pResp.ok ? await pResp.json() : { value: [] };
      const positions = (pData.value || []).map((p: any) => {
        const tVals = p.FieldValuesAsText || {};
        let text = tVals.Name || tVals.Title || p.Name || p.Title || '';
        if (!text) {
           const pKey = Object.keys(tVals).find(k => k.toLowerCase().includes('name') || k.toLowerCase().includes('title'));
           if (pKey) text = tVals[pKey];
        }
        return { key: String(p.Id), text: text || `PosID:${p.Id}` };
      });

      // 2. Fetch Business Profiles
      let bUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Business Profiles')/items?$top=5000&$expand=FieldValuesAsText`;
      let bResp = await fetch(bUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      const bData = bResp.ok ? await bResp.json() : { value: [] };
      const businesses = (bData.value || []).map((b: any) => {
        const tVals = b.FieldValuesAsText || {};
        let text = tVals.Business_x0020_Profile || tVals.BusinessProfile || tVals.Title || b.Title || '';
        if (!text) {
           const bKey = Object.keys(tVals).find(k => k.toLowerCase().includes('business') && k.toLowerCase().includes('profile'));
           if (bKey) text = tVals[bKey];
        }
        return { key: String(b.Id), text: text || `BizID:${b.Id}` };
      });

      // 3. Fetch Employees
      const eUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Employees')/items?$top=5000&$expand=FieldValuesAsText`;
      const eResp = await fetch(eUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!eResp.ok) return [];
      const eData = await eResp.json();

      return (eData.value || []).map((e: any) => {
        const tVals = e.FieldValuesAsText || {};

        let empName = String(tVals.Employee_x0020_Name || tVals.EmployeeName || e.Employee_x0020_Name || e.EmployeeName || tVals.Title || e.Title || '').trim();
        if (!empName || empName === 'undefined' || empName.includes(';#')) {
            const fKey = Object.keys(tVals).find(k => k.toLowerCase().includes('first') && k.toLowerCase().includes('name'));
            const lKey = Object.keys(tVals).find(k => k.toLowerCase().includes('last') && k.toLowerCase().includes('name'));
            
            const fn = fKey ? String(tVals[fKey]).trim() : String(e.FirstName || e.First_x0020_Name || '').trim();
            const ln = lKey ? String(tVals[lKey]).trim() : String(e.LastName || e.Last_x0020_Name || '').trim();
            
            if (fn || ln) {
                empName = [fn, ln].filter(Boolean).filter(x => x !== 'undefined' && x !== 'null').join(' ');
            }
        }
        if (!empName || empName === 'undefined') empName = `(Missing Name ID:${e.Id})`;

        // Position Lookup
        let rawPos = String(e.PositionId || e.Position_x0020_TypeId || tVals.Position || tVals.Position_x0020_Type || '').trim();
        if (!rawPos || rawPos === 'undefined') {
            const pKey = Object.keys(e).find(k => k.toLowerCase().includes('position') && k.toLowerCase().endsWith('id'));
            if (pKey) rawPos = String(e[pKey]).trim();
        }
        let finalPos = rawPos;
        const matchedPos = positions.find((x: any) => x.key === rawPos);
        if (matchedPos && !matchedPos.text.startsWith('PosID:')) finalPos = matchedPos.text;

        // Business Profile Lookup
        let rawBiz = String(e.Business_x0020_ProfileId || e.BusinessProfileId || e.BusinessProfileId || tVals.Business_x0020_Profile || tVals.BusinessProfile || '').trim();
        if (!rawBiz || rawBiz === 'undefined') {
            const bKey = Object.keys(e).find(k => k.toLowerCase().includes('business') && k.toLowerCase().includes('profile') && k.toLowerCase().endsWith('id'));
            if (bKey) rawBiz = String(e[bKey]).trim();
        }
        let finalBiz = rawBiz;
        const matchedBiz = businesses.find((x: any) => x.key === rawBiz);
        if (matchedBiz && !matchedBiz.text.startsWith('BizID:')) finalBiz = matchedBiz.text;

        // Main Phone
        let phone = String(tVals.Mobile_x0020_Phone || tVals.MobilePhone || tVals.MainPhone || e.MobilePhone || '').trim();
        if (!phone || phone === 'undefined') {
            const mKey = Object.keys(tVals).find(k => k.toLowerCase().includes('mobile') || k.toLowerCase().includes('phone'));
            if (mKey) phone = String(tVals[mKey]).trim();
        }

        // Email
        let email = String(tVals.Email || e.Email || '').trim();
        if (!email || email === 'undefined') {
            const emKey = Object.keys(tVals).find(k => k.toLowerCase().includes('email'));
            if (emKey) email = String(tVals[emKey]).trim();
        }

        const strParts = [empName, finalPos, finalBiz, phone, email]
            .map(x => String(x).trim())
            .filter(Boolean)
            .filter(x => x !== 'undefined' && x !== 'null');
            
        return { key: e.Id || e.ID, text: strParts.join(' | ') };
      });
    } catch (e: any) {
      return [{ key: -99, text: `ERR: ${(e as Error).message}` }];
    }
  }

  private async fetchLookup(listName: string, fieldName: string): Promise<ILookupOption[]> {
    try {
      let url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$top=5000&$expand=FieldValuesAsText`;
      let resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
      
      // If the list name with spaces doesn't work, try it without spaces
      if (!resp.ok) {
        url = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName.replace(/\s/g, ''))}')/items?$top=5000&$expand=FieldValuesAsText`;
        resp = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
      }

      if (!resp.ok) return [];
      const data = await resp.json();
      return (data.value || [])
        .map((item: any) => {
          const tVals = item.FieldValuesAsText || {};
          let text = tVals[fieldName] || item[fieldName] || tVals.Title || tVals.Name || item.Title || item.Name;
          
          if (!text) {
             const cleanField = fieldName.toLowerCase().replace(/_x0020_/g, '');
             const possibleKey = Object.keys(tVals).find(k => 
                k.toLowerCase().includes(cleanField) || 
                k.toLowerCase().includes('profile') ||
                k.toLowerCase().includes('name') ||
                k.toLowerCase().includes('title')
             );
             text = possibleKey ? tVals[possibleKey] : `Item ${item.Id}`;
          }

          return { key: item.Id, text: text as string };
        });
    } catch {
      return [];
    }
  }

  private async fetchWorkers(_subcontractors: ILookupOption[]): Promise<ILookupOption[]> {
    try {
      // 1. Fetch Positions explicitly to map the ID (e.g. 349) to the actual string text
      let pUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Position Types')/items?$top=5000&$expand=FieldValuesAsText`;
      let pResp = await fetch(pUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!pResp.ok) {
        pUrl = `${this.siteUrl}/_api/web/lists/getbytitle('PositionTypes')/items?$top=5000&$expand=FieldValuesAsText`;
        pResp = await fetch(pUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      }
      if (!pResp.ok) {
        pUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Positions')/items?$top=5000&$expand=FieldValuesAsText`;
        pResp = await fetch(pUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      }
      
      const pData = pResp.ok ? await pResp.json() : { value: [] };
      const positionsFound = pResp.ok;
      
      const positions = (pData.value || []).map((p: any) => {
        const tVals = p.FieldValuesAsText || {};
        let text = tVals.Name || tVals.Title || tVals.Position || p.Name || p.Title || p.Position;
        
        if (!text) {
          const possibleKey = Object.keys(tVals).find(k => 
              k.toLowerCase().includes('name') || 
              k.toLowerCase().includes('title') || 
              k.toLowerCase().includes('position') ||
              k.toLowerCase().includes('type')
          );
          text = possibleKey ? tVals[possibleKey] : `Keys:${Object.keys(tVals).slice(0, 3).join('_')}`;
        }
        return { key: String(p.Id), text: text || `PosID:${p.Id}` };
      });

      // 2. Fetch Subcontractors explicitly to map the ID (e.g. 2) to the actual string text
      const sUrl = `${this.siteUrl}/_api/web/lists/getbytitle('Subcontractors')/items?$top=5000&$expand=FieldValuesAsText`;
      const sResp = await fetch(sUrl, { headers: { Accept: 'application/json;odata=nometadata' } });
      const sData = sResp.ok ? await sResp.json() : { value: [] };
      const subconts = (sData.value || []).map((s: any) => {
        const tVals = s.FieldValuesAsText || {};
        const text = tVals.Company_x0020_Name || tVals.CompanyName || tVals.Title || tVals.Company || s.Title || `CompID:${s.Id}`;
        return { key: String(s.Id), text };
      });

      // 3. Fetch Contacts
      const tryUrls = [
        `${this.siteUrl}/_api/web/lists/getbytitle('Contacts')/items?$top=5000&$expand=FieldValuesAsText`,
        `${this.siteUrl}/_api/web/lists/getbytitle('Contact')/items?$top=5000&$expand=FieldValuesAsText`
      ];
      let cResp: Response | undefined;
      for (const url of tryUrls) {
        const r = await fetch(url, { headers: { Accept: 'application/json;odata=nometadata' } });
        if (r.ok) { cResp = r; break; }
      }
      if (!cResp || !cResp.ok) return [{ key: -1, text: 'ERR: Contacts list not found' }];

      const cData = await cResp.json();
      return (cData.value || []).map((c: any) => {
        const txt = c.FieldValuesAsText || {};

        // 1. Employee Name
        const empName = txt.Employee_x0020_Name || txt.EmployeeName || c.Employee_x0020_Name || c.Title || c.Name || 'Unknown';

        // 2. Extract Raw ID for Position (usually resolves currently as "349")
        const rawPosId = String(txt.Position || txt.Position_x0020_Type || c.PositionId || '').trim();
        let posText = positionsFound ? rawPosId : `${rawPosId} (List Not Found)`;
        const matchedPos = positions.find((x: any) => x.key === rawPosId);
        if (matchedPos && !matchedPos.text.startsWith('PosID:')) posText = matchedPos.text;

        // 3. Extract Raw ID for Company (usually resolves currently as "2")
        const rawCompId = String(txt.Company || txt.Company_x0020_Name || c.CompanyId || c.SubcontractorId || '').trim();
        let compText = rawCompId;
        const matchedComp = subconts.find((x: any) => x.key === rawCompId);
        if (matchedComp && !matchedComp.text.startsWith('CompID:')) compText = matchedComp.text;

        // 4. Join components and clear out empty strings perfectly
        const parts = [empName, posText, compText, 'Supply Workforce']
          .map(s => String(s).trim())
          .filter(s => s.length > 0);

        return { key: c.Id, text: parts.join(' | ') };
      });
    } catch (e: any) {
      return [{ key: -99, text: `ERR: ${(e as Error).message}` }];
    }
  }
}

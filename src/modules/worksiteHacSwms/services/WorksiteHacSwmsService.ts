import { WebPartContext } from '@microsoft/sp-webpart-base';
import { BaseSPService } from '../../../common/services/BaseSPService';
import {
    defaultWorksiteHacSwmsListConfig,
    IWorksiteHacSwmsListConfig
} from '../WorksiteHacSwmsConfig';

export class WorksiteHacSwmsService extends BaseSPService {
    private readonly _listConfig: IWorksiteHacSwmsListConfig;

    constructor(context: WebPartContext, listConfig: IWorksiteHacSwmsListConfig = defaultWorksiteHacSwmsListConfig) {
        super(context);
        this._listConfig = listConfig;
    }

    public getListConfig(): IWorksiteHacSwmsListConfig {
        return this._listConfig;
    }
}

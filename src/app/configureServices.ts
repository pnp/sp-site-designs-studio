import { ServiceScope, Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SiteDesignsServiceKey } from "../services/siteDesigns/SiteDesignsService";
import { MockSiteDesignsService } from "../services/siteDesigns/SiteDesignsMockService";
import { SiteScriptSchemaServiceKey } from "../services/siteScriptSchema/SiteScriptSchemaService";
import { RenderingServiceKey } from "../services/rendering/RenderingService";
import { hubSitePickerRenderer } from "../components/propertyInputRenderers/HubSitePicker";
import { appPickerRenderer } from "../components/propertyInputRenderers/AppPicker";
import { themePickerRenderer } from "../components/propertyInputRenderers/ThemePicker";
import { listTemplatePickerRenderer } from "../components/propertyInputRenderers/ListTemplatePicker";
import { SiteScriptSamplesServiceKey } from "../services/siteScriptSamples/SiteScriptSamplesService";


async function configureTestServices(webPartContext: WebPartContext): Promise<ServiceScope> {
    const childScope = webPartContext.serviceScope.startNewChild();

    childScope.createAndProvide(SiteDesignsServiceKey, MockSiteDesignsService);

    childScope.finish();

    await new Promise<void>((resolve, reject) => {
        childScope.whenFinished(() => {
            try {
                const siteScriptSchema = childScope.consume(SiteScriptSchemaServiceKey);
                const siteScriptSchemaConfigPromise = siteScriptSchema.configure();

                Promise.all([siteScriptSchemaConfigPromise]).then(() => {
                    resolve();
                }).catch(error => {
                    reject(error);
                });
            } catch (error) {
                reject(error);
            }
        });
    });

    return childScope;
}

async function configureProdServices(webPartContext: WebPartContext): Promise<ServiceScope> {
    const childScope = webPartContext.serviceScope.startNewChild();
    // TODO Create and configure custom service instances here

    childScope.finish();

    await new Promise<void>((resolve, reject) => {
        childScope.whenFinished(() => {
            try {
                const siteScriptSample = childScope.consume(SiteScriptSamplesServiceKey);
                siteScriptSample["_availableRepositories"] = [
                    {
                        key: 'PnP',
                        owner: 'pnp',
                        repository: 'sp-dev-site-scripts',
                        branch: 'master',
                        samplesFolderPath: 'samples'
                    }
                ];

                const siteScriptSchema = childScope.consume(SiteScriptSchemaServiceKey);
                const siteScriptSchemaConfigPromise = siteScriptSchema.configure();

                Promise.all([siteScriptSchemaConfigPromise]).then(() => {
                    // The schema service must be configured before the customization config can be done
                    const rendering = childScope.consume(RenderingServiceKey);
                    rendering.customizeActionPropertyRendering("joinHubSite", null, "hubSiteId", hubSitePickerRenderer);
                    rendering.customizeActionPropertyRendering("installSolution", null, "id", appPickerRenderer);
                    rendering.customizeActionPropertyRendering("applyTheme", null, "themeName", themePickerRenderer);
                    rendering.customizeActionPropertyRendering("createSPList", null, "templateType", listTemplatePickerRenderer);
                    resolve();
                }).catch(error => {
                    reject(error);
                });
            } catch (error) {
                reject(error);
            }
        });
    });

    return childScope;
}

export default function configureServices(webPartContext: WebPartContext): Promise<ServiceScope> {
    switch (Environment.type) {
        case EnvironmentType.Local:
        case EnvironmentType.Test:
            return configureTestServices(webPartContext);
        default:
            return configureProdServices(webPartContext);
    }
} 
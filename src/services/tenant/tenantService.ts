import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { AadTokenProviderFactory } from "@microsoft/sp-http";
import { graph } from "@pnp/graph/presets/all";
import { sp } from "@pnp/sp";
import { IPrincipal } from "../../models/IPrincipal";
import { GraphBatch } from "@pnp/graph/batch";


export interface ITenantService {
    /**
     * Search users or mail enabled groups
     * @param criteria the criteria to search for
     * @returns Promise<IPrincipal[]> The list of principal matching the criteria
     */
    searchPrincipals(criteria: string): Promise<IPrincipal[]>;
    ensurePrincipalInfo(principals: IPrincipal[]): Promise<IPrincipal[]>;
}

class TenantService implements ITenantService {

    private pageContext: PageContext = null;
    constructor(serviceScope: ServiceScope) {

        serviceScope.whenFinished(() => {
            this.pageContext = serviceScope.consume(PageContext.serviceKey);
            const tokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);
            graph.setup({
                spfxContext: {
                    aadTokenProviderFactory: tokenProviderFactory,
                    pageContext: this.pageContext
                }
            });
        });
    }

    public async ensurePrincipalInfo(principals: IPrincipal[]): Promise<IPrincipal[]> {

        const ensuredPrincipals = principals.map(p => ({ ...p }));

        const promises = [];
        ensuredPrincipals.forEach(p => {

            // Don't try to ensure external users
            // Graph won't allow querying external users by login name
            if (p.principalName.indexOf("#EXT#") >= 0) {
                console.log(`Will skip external user ${p.principalName}`);
                return;
            }

            // If the type of the principal is already known, no need to ensure
            if (p.type) {
                return;
            }
 
            // If the login name contains an @, it is most likely a user principal name
            const isLoginName = p.principalName.indexOf('@') >= 0;
            const mailNickname = isLoginName ? p.principalName.split('@')[0] : '';
            // If not a login name, probably a group's Id
            const groupId = !isLoginName ? p.principalName : '';

            if (isLoginName) {
                promises.push(graph.users
                    .filter(`userPrincipalName eq '${p.principalName}'`)
                    .select("userPrincipalName", "mail", "id", "displayName")
                    .top(1)
                    .get().then(res => {
                        if (res.length == 1) {
                            p.displayName = res[0].displayName;
                            p.id = res[0].id;
                            p.type = "user";
                        } else if (res.length == 0) {
                            console.warn(`User not found with principal name ${p.principalName}`);
                        }
                    }));
            }

            let groupsQueryFilter = 'mailEnabled eq true';
            if (mailNickname || groupId) {
                groupsQueryFilter += ' and (';
                if (mailNickname) {
                    groupsQueryFilter += `mailNickname eq '${mailNickname}'`;
                    if (groupId) {
                        groupsQueryFilter += ' or ';
                    }
                }
                if (groupId) {
                    groupsQueryFilter += `id eq '${groupId}'`;
                }
                groupsQueryFilter += ')';
            }
            console.log("Groups query filter: ", groupsQueryFilter);
            promises.push(graph.groups
                .filter(groupsQueryFilter)
                .select("mailNickname", "mail", "id", "displayName")
                .top(1)
                .get()
                .then(res => {
                    if (res.length == 1) {
                        p.displayName = res[0].displayName;
                        p.id = res[0].id;
                        p.type = "group";
                    } else if (res.length == 0) {
                        console.warn(`Group not found with mailNickname matching ${mailNickname}`);
                    }
                }));
        });

        await Promise.all(promises);
        return ensuredPrincipals;
    }



    public async searchPrincipals(criteria: string): Promise<IPrincipal[]> {
        const ensuredCriteria = escape(criteria);
        // TODO use $batch to improve this
        const matchingGroups = await graph.groups
            .filter(`mailEnabled eq true and (startswith(mailNickname,'${ensuredCriteria}') or startswith(displayName,'${ensuredCriteria}'))`)
            .select("mailNickname", "mail", "id", "displayName")
            .top(15)
            .get();
        const matchingUsers = await graph.users
            .filter(`startswith(userPrincipalName,'${ensuredCriteria}') or startswith(displayName,'${ensuredCriteria}')`)
            .select("userPrincipalName", "mail", "id", "displayName")
            .top(15)
            .get();

        const result = matchingGroups.map(g => ({ displayName: g.displayName, id: g.id, principalName: g.mail, type: "group" } as IPrincipal))
            .concat(...matchingUsers.map(u => ({ displayName: u.displayName, id: u.id, principalName: u.userPrincipalName, type: "user" } as IPrincipal)));

        return result;
    }

}

export const TenantServiceKey = ServiceKey.create<ITenantService>(
    'YPCODE:SDSv2:TenantService',
    TenantService
);

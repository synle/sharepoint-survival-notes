# sharepoint-survival-notes

## Code Related

### Requirements

```
Node:
v10.24.1

Npm:
6.14.12
```

### typescript config note:

For import alias shorten path, edit `tsconfig.json`

```
{
  "compilerOptions": {
    ...
    "baseUrl": "./src",
    ...
  }
}
```

### Sample Scripts

```
"scripts": {
    "format": "sh format.sh",
    "build": "gulp bundle",
    "clean": "gulp clean",
    "test": "gulp test",
    "start": "Node_NO_HTTP2=1 gulp serve",
    "serve": "Node_NO_HTTP2=1 gulp serve"
  },
```

### Formating script

```
npx prettier --write ./src/**/**/**/**/**/*.ts && \
npx prettier --write ./src/**/**/**/**/**/*.tsx && \
npx prettier --write ./src/**/**/**/**/**/*.scss && \
npx prettier --write --parser json **/**/**/*.json .prettierrc && \
echo 'Done formatting'
```

### Build & Package & Deploy

More info on this here: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/hosting-webpart-from-office-365-cdn

```
echo '''
========================================================
= gulp clean
========================================================
'''
./node_modules/.bin/gulp clean


echo '''
========================================================
= gulp bundle
========================================================
'''
./node_modules/.bin/gulp bundle --ship

echo '''
========================================================
= gulp package-solution
========================================================
'''
./node_modules/.bin/gulp package-solution --ship

echo '''
========================================================
= build output
========================================================
'''
echo sharepoint/solution/*.sppkg


```

### Testing webparts live on a site

Append `/_layouts/15/workbench.aspx` to the URL. Something that looks like this

```
https://mytestorg.sharepoint.com/sites/MyTestSite/_layouts/15/workbench.aspx
```

### Useful links:

- Webpart training: https://www.youtube.com/watch?v=yc1IYgYp7qQ&list=PLR9nK3mnD-OXvSWvS2zglCzz4iplhVrKq

#### Other useful Links:

- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
- https://www.c-sharpcorner.com/article/how-to-create-sharepoint-frameworkspfx-extensions/
- https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/list-form-configuration
- https://collab365.community/publish-app-to-non-developer-site-collection/
- https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-started-creating-sharepoint-hosted-sharepoint-add-ins
- https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant
- https://www.youtube.com/watch?v=yc1IYgYp7qQ

#### Common PNP Webpart projects

https://github.com/pnp/sp-dev-solutions

### How to access dev app catalog and upload your plugin?

- Click on gear icon top right
- Choose `Site Content` under Settings > Sharepoint
- Under the content grid, Click on `App for Sharepoint`
- There you can upload your sharepoint package (sppkg).
- After you upload and deploy, you need to add the addon to your site.

## Sharepoint config related

### Enable Content type

https://support.microsoft.com/en-us/office/add-a-content-type-to-a-list-or-library-917366ae-f7a2-47ad-87a5-9689a1884e60

Under Content Types, select Add from existing site content types. If Content Types doesn't appear, select Advanced settings, and select Yes under Allow management of content types?, and then select OK.

## Flow Related

### API Request Headers

This is a common API request header for the API

```
content-type
   application/json;odata=verbose
Accept
  application/json;odata=verbose
```

### Open flow from from List View

On the top bar > choose `Integrate` > choose `Power Automate` > Choose either `Create a flow` or `See your flows`

### Flows Actions

- Delay action for delaying
- Control / Condition
- Send email notification V3 (make sure you type it in, not copy and paste)

### Complete guide on how to use flow to publish page

https://techcommunity.microsoft.com/t5/power-apps-power-automate/create-site-page-with-flow/m-p/885468
http://www.sites.se/2018/08/sharepoint-modern-pages-microsoft-flow/

### Get file by server URL

https://mytestorg.sharepoint.com/sites/MyTestSite/_api/web/GetFileByServerRelativeUrl('/sites/MyTestSite/SitePages/sy-default-page-template.aspx')

### Flow get users members in a group

https://www.c-sharpcorner.com/article/send-mail-to-sharepoint-group-members-using-power-automate/

```
fetch("https://mytestorg.sharepoint.com/sites/MyTestSite/_api/web/sitegroups/getbyname('LMS%20Admin')/users", {
  "headers": {
     "Accept": "application/json; odata=verbose",
      "content-type": "application/json; odata=verbose",
  }
}).then(async (res) => console.log(JSON.stringify(await res.json())));
```

```
body('Parse_JSON')?['d']?['results']
```

```
_api/web/sitegroups/getbyname('MYGroup')/users
```

### Add comments in flow

https://laptrinhx.com/add-comments-to-sharepoint-list-items-using-the-rest-api-3384205990/

\_api/web/lists/GetByTitle('Asset Requests')/items(@{triggerOutputs()?['body/ID']})/Comments()

```
{
    "inputs": {
        "host": {
            "connectionName": "shared_sharepointonline",
            "operationId": "HttpRequest",
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
        },
        "parameters": {
            "dataset": "https://mytestorg.sharepoint.com/sites/MyTestSite",
            "parameters/method": "POST",
            "parameters/uri": "_api/web/lists/GetByTitle('Asset Requests')/items(@{triggerOutputs()?['body/ID']})/Comments()",
            "parameters/headers": {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;charset=UTF-8"
            },
            "parameters/body": "{\"text\": \"Add a new comment 123 - @{triggerOutputs()?['body/Author/DisplayName']} @{triggerOutputs()?['body/Author/Email']}\"}"
        },
        "authentication": "@parameters('$authentication')"
    }
}
```

### Copy a page (template) to a new page

```
_api/web/GetFileByServerRelativeUrl('/sites/MyTestSite/SitePages/sy-default-page-template.aspx')/copyTo('/sites/MyTestSite/SitePages/@{variables('Page Name')}')
```

### Use flow action `Get File MetaData` to get the pageId

/SitePages/@{variables('Page Name')}

We might need to escape the string with `%252f`

```
SitePages%252f@{variables('Page Name')}
```

### Use flow action `Update File Property` to update metadata

Action `Update file property`

### Checkout pages for Edit

POST \_api/sitepages/pages({id})/CheckOutPage
POST \_api/sitepages/pages(@{variables('Page ID')})/CheckOutPage

### Save pages as Draft

POST \_api/sitepages/pages({id})/SavePageAsDraft
POST \_api/sitepages/pages(@{variables('Page ID')})/SavePageAsDraft

### Publish pages

POST \_api/sitepages/pages(20)/publish
POST \_api/sitepages/pages(@{outputs('Get_file_metadata')?['body/ItemId']})/publish

{"\_\_metadata":{"type":"SP.Publishing.SitePage"}}

## Search Schema

### Microsoft KQL (Keyword Query Language) Syntax

https://docs.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference

#### Exclude dispForm using KQL

https://sharepoint.stackexchange.com/questions/246565/exclude-forms-dispform-aspx-from-search-results-search-api-sharepoint-online

### Refiner

https://docs.microsoft.com/en-us/sharepoint/manage-search-schema#refine-on-managed-properties

```
If you want to use a managed property as a refiner on the search results page, use the setting refinable. This setting is only available on the built-in managed properties, and only affects the classic search experience. If you need to use a new managed property, or an auto-generated managed property, as a refiner, rename an existing, unused managed property (that's refinable) by using an alias. There are quite a few managed properties available for this purpose. They have names such as "RefinableString00" and "RefinableDate19."

For example, you create a new site column called "NewColors", and you want users to be able to use "NewColors" as an option when they refine on the search results. In the search schema, you choose an unused managed property, for example "RefinableString00", and rename the property to "NewColors" by using the Alias setting. Then, you map this new managed property to the relevant crawled property
```

### Manual Reindex Sharepoint Search

https://social.technet.microsoft.com/Forums/en-US/5f1e9084-0ead-4dd8-b182-3ddd704679ff/refinablestring-actual-values-are-not-shown-in-the-refinement-panel?forum=onlineservicessharepoint

Go to gear icon for setting in the site->site settings->Search-> Search and offline availability->click “Reindex site” button. Check if you can see the value in the refinement.

## Setting up PNP webpart refiner

PNP Refiner youtube training video : https://youtu.be/y760Okrz-Oo

## Sample Scripts

### Sample Script to get digest and make API calls.

```

function _getBaseUrl() {
  const baseUrl = location.href.match(
    /https:\/\/[a-z0-9._]+\/(teams|sites)\/[a-z0-9_]+\//gi,
  )[0];
  return baseUrl;
}

function _fetch<T>(url, props, digest?: string): Promise<T> {
  const baseUrl = _getBaseUrl();
  url = `${baseUrl}${url}`;
  return fetch(url, {
    ...props,
    headers: {
      accept: 'application/json; odata=nometadata',
      'content-type': 'application/json;odata=verbose;charset=utf-8',
      'x-requestdigest': digest,
    },
    referrer: baseUrl,
  }).then((r) => r.json());
}

async function _fetchSecure<T>(url, props) {
  const digest = await _getDigest();
  return _fetch<T>(url, props, digest);
}

/**
 * This API gets a digest (similar to CSRF token) used to
 * make further API calls...
 * @returns
 */
function _getDigest() {
  return _fetch<string>('/_api/contextinfo', {
    method: 'POST',
  }).then((r: any) => r.FormDigestValue);
}

/**
 * Core API used to search data from sharepoint
 *
 * More info on this here: https://docs.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview
 * @param config
 * @returns
 */
export function postQuery(config: any) {
  return _fetchSecure<SharepointApiResponse.PostQueryResponse>(
    '/_api/search/postquery',
    {
      method: 'POST',
      body: JSON.stringify(config),
    },
  );
}

```

## Powerapp Form Notes

### Open powerapp from from List View

Edit a record > Then choose More option > Customize with Power Apps

### Simple if else

```
If("Pending" in approvalDropdown.SelectedItems.Value, true)
```

### Display Mode validation check

Set `DisplayMode`

https://www.chakkaradeep.com/2018/01/26/show-and-hide-controls-in-sharepoint-list-custom-forms-based-on-sharepoint-form-modes-in-powerapps/
https://www.youtube.com/watch?v=9gI9OscTLD0

```
if(SomeForm.Valid, DisplayMode.Edit, Disabled)
```

```
If("Pending" in approvalDropdown.SelectedItems.Value && SharePointForm1.Mode = FormMode.Edit, true, false)

0 – Edit Mode
1 – New Mode
2 – Display Mode
```

### Submit form and patch data

https://docs.microsoft.com/en-us/powerapps/maker/canvas-apps/functions/function-patch

### Patch User Claims

https://sharepoint.stackexchange.com/questions/276802/how-to-update-patch-sharepoint-group-in-person-and-group-field-using-powerapps

### Patch choice (single select field)

https://baizini-it.com/blog/index.php/2018/01/15/powerapps-how-to-update-sharepoint-choice-and-lookup-type-columns-with-patch/

### More complete option

```
Set(
    ApprovedByUser,
    {
        Claims: User().FullName,
        Department: "",
        DisplayName: User().FullName,
        Email: User().Email,
        JobTitle: "",
        Picture: ""
    }
);

Patch(
    'Asset Requests',
    SharePointIntegration.Selected,
    {
        description: "Approved by " & " " & User().Email,
        'approved by': ApprovedByUser,
        approval:  {
            Value: "Approved",
            '@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference"
        }
    }
);

Notify("Done approved this asset", Success);

Refresh(SharePointIntegration.Selected);
```

### Close form

```
Exit();
```

### Direct link to approval app

https://australia.flow.microsoft.com/manage/approvals/received
https://flow.microsoft.com/manage/approvals/received
https://joannecklein.com/2018/11/30/using-microsoft-flow-to-know-when-a-modern-page-is-published/

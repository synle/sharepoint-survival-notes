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
npx prettier --write src/**/*.{ts,tsx,scss} && \
npx prettier --write *.md && \
npx prettier --write --parser json *.json .prettierrc && \
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

### Flow actions:

- `Parallel action`: can be used to handle reminders. Will likely need to be combined with `Delay` and `Delay Until`. For example, remind user to approve a task after 1 week or 1 month, etc...
- `Scope action`: can be used to group a list of related actions / items into a single container, super useful for copy and paste or simply move actions around.
- `Apply to Each`: used for loop, is used in place of multi select fields as a way to return multi select value
- `Variable`: use variable as constants or a local buffer for your flow. Must be defined on the top scope with `initialize variable`. Then `set variable` can be used when you want to assign a value to it at a later time. 
-  Copy and Paste / Clipboard: click on ellipsis and then choose copy, create new action and choose `Clipboard`, and apply the action


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

### Child Flow using HTTP Incoming Request
Webhook can be used as a way to manage child flows

1. For your child flow - use trigger to be called "When a HTTP request is received". You need a schema - you can choose generate schema from sample data. The Flow app will parse and generate the schema.

Mine looks like this
Sample data:
```
{
  "asset_request_id": "54",
  "published_page_id": "1033"
}
```

Schema:
```
{
    "type": "object",
    "properties": {
        "asset_request_id": {
            "type": "string"
        },
        "published_page_id": {
            "type": "string"
        }
    }
}
```

2. At the end you can choose to respond the http - here I returns 200 OK
3. After you save - the child flow will generate a webhook within azure, you will need this URL to trigger it from the parent flow.
4. My parent flow is simple, just need to make an HTTP action with POST to the URL provided by the child flow.



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

#### The plain old fetch API call
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
      'content-type': 'application/json;odata=verbose;',
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
export function postQuery(config: unknown) {
  return _fetchSecure<PostQueryResponse>(
    '/_api/search/postquery',
    {
      method: 'POST',
      body: JSON.stringify(config),
    },
  );
}

```

#### The sharepoint way using SPHttpClient

Refer to these links for more details:

- https://github.com/pnp/sp-dev-fx-webparts/blob/main/samples/sharepoint-crud/src/webparts/reactCrud/ReactCrudWebPart.ts
- https://www.vrdmn.com/2017/01/working-with-rest-api-in-sharepoint.html


##### Install the package `@microsoft/sp-http`

```
npm i --save @microsoft/sp-http
```

Then import it

```
import {
  ODataVersion,
  SPHttpClient,
  SPHttpClientConfiguration,
} from '@microsoft/sp-http';
```

##### Pass in the `absoluteUrl` and `spHttpClient` into your subcomponent

This method is preferred because we pull these 2 items from the context instead of using regex for `absoluteUrl` or plain vanialla `fetch` in js

###### The interace for props
```
export interface AwesomeWebpartProps {
  siteUrl: string;
  spHttpClient: SPHttpClient;
}
```

###### The base webpart
```
export default class AwesomeWebpart extends BaseClientSideWebPart<AwesomeWebpartProps> {
  public render(): void {
    const element = React.createElement(AwesomeComponent, {
      ...{
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
      },
      ...this.properties,
    });

    ReactDom.render(element, this.domElement);
  }
  ...
}
```

##### The actual component

```
interface AwesomeComponentStates{
  // add your states
}

interface PostQueryResponse{
  // add your response type here
}

// rest client config
// Since the SP Search REST API works with ODataVersion 3, we have to create a new ISPHttpClientConfiguration object with defaultODataVersion = ODataVersion.v3
// Override the default ODataVersion.v4 flag with the ODataVersion.v3\
//
// refer to this link for more note: https://www.vrdmn.com/2017/01/working-with-rest-api-in-sharepoint.html
const clientConfigODataV3: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith(
  {
    defaultODataVersion: ODataVersion.v3,
  },
);

export default class AwesomeComponent extends React.Component<
  AwesomeComponentProps,
  AwesomeComponentStates
> {
  constructor(props) {
    super(props);

    // default state
    this.state = {};
  }

  componentDidMount() {
    this.doSearchQuery();
  }
  
  private async doSearchQuery(
    config: unknown,
  ): Promise<PostQueryResponse> {
    return this.props.spHttpClient
      .fetch(
        `${this.props.siteUrl}/_api/search/postquery`,
        clientConfigODataV3,
        {
          headers: {
            accept: 'application/json; odata=nometadata',
            'content-type': 'application/json;odata=verbose;',
          },
          method: 'POST',
          body: JSON.stringify(config),
        },
      )
      .then((r) => r.json());
  }
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

### New Form and Show Data After Submission
#### App
##### App - OnStart
```
ResetForm(SharepointSubmissionForm);
NewForm(SharepointSubmissionForm);
```

#### Save Button
##### Save Button - OnSelect
```
SubmitForm(SharepointSubmissionForm);
```


#### Form
##### Form - OnSuccess
```
Set(varLastSubmit, SharepointSubmissionForm.LastSubmit);

ViewForm(SharepointSubmissionForm);
SharepointSubmissionForm.Mode = FormMode.View;


Notify(Concatenate("Data saved successfully: ", Text(SharepointSubmissionForm.LastSubmit.ID)), NotificationType.Success)
```

### Filter out options in Dropdown
#### Dropdown - Items
```
Filter(Choices([@'Asset Submission'].LDH_AssetStatus), Value <> "Archived")
```


### Changing Render Type (Control Type) of Sharepoint Form
- Go to the sharepoint form
- Under `Properties` Tab, choose `Edit Fields`
![image](https://user-images.githubusercontent.com/3792401/124297484-f79a2980-db0f-11eb-9e45-5d4487dc1fce.png)

- Click on control type and modify it.
![image](https://user-images.githubusercontent.com/3792401/124297524-041e8200-db10-11eb-9e6e-7100cc009230.png)



### Stretching the Canvas
Go to Advanced tab and update Width and Height
```
Max(App.Width, App.MinScreenWidth)
Max(App.Height, App.MinScreenHeight)
```

### Direct link to approval app

https://australia.flow.microsoft.com/manage/approvals/received
https://flow.microsoft.com/manage/approvals/received
https://joannecklein.com/2018/11/30/using-microsoft-flow-to-know-when-a-modern-page-is-published/



## Other tips and tricks
### Script to get the name and column from site content types and other things
Run this in the console
```
data = [...document.querySelectorAll('#columnstable a')].map(s => {
  try{
    return `|${s.innerText}|${decodeURIComponent(s.href).match(/Field=[a-zA-Z0-9_]+/)[0].replace('Field=', '')}|`
  }  catch(err){
    return '';
  }
})
copy(`|Display|Column Name|\n|--------|--------|\n` + data.filter(s => !!s).join('\n'))
```


### Definitions for the naming of sharepoint crawled properties
https://sharepoint.stackexchange.com/questions/227344/where-is-the-definitive-list-of-crawled-properties-in-sharepoint-2016

### Sharepoint managed properties name:

https://sharepointgypsy.wordpress.com/2019/04/22/sharepoint-2016-finding-crawled-properties-and-manage-properties-by-name/

- `DATE` : Date and Time
- `TEXT` : Single Line of Text
- `MTXT` : Multiple Lines of Text
- `CHCS` : Choice
- `CHCM` : Choice that Allows Multiple Choice
- `NMBR` : Number
- `CURR` : Currency
- `BOOL` : Yes/No
- `USER` : Person or Group
- `URLH` : Hyperlink or Picture
- `HTML` : Publishing HTML
- `IMGE` : Publishing Image
- `LINK` : Publishing Link
- `INTG` : Integer
- `GUID` : GUID
- `CHCG` : Grid Choice
- `CTID` : Content Type ID Field
- `RAVG` : SPS Average Rating
- `RCNT` : SPS Rating Count


### How to parse RefinableString of type People

Sample string
```
jdoe@sharepoint-test.com | John Doe | 693X30232X667X6X656X626572736869707X677473756X616861406X696X6X6564696X2X62697X i:0#.f|membership|jdoe@sharepoint-test.com;bswagger@sharepoint-test.com | Bob Swagger | 693X30232X667X6X656X626572736869707X6X656X6565406X696X6X6564696X2X62697X i:0#.f|membership|bswagger@sharepoint-test.com
```

Sample method used to extract the people
```
export function getParsedPeopleField(personListString: string): PeopleEmail[] {
  try {
    return personListString
      .split(';')
      .map((s1) => s1.split(' | ').map((s2) => s2.trim()))
      .map(([email, fullName]) => ({
        email,
        fullName,
        profileImageSrc: `/_layouts/15/userphoto.aspx?size=L&username=${email}`
      }));
  } catch (err) {
    return [];
  }
}
```

Sample response from the above method
```
[
    {
        "email": "jdoe@sharepoint-test.com",
        "fullName": "John Doe",
        "profileImageSrc": "/_layouts/15/userphoto.aspx?size=L&username=jdoe@sharepoint-test.com"
    },
    {
        "email": "bswagger@sharepoint-test.com",
        "fullName": "Bob Swagger",
        "profileImageSrc": "/_layouts/15/userphoto.aspx?size=L&username=bswagger@sharepoint-test.com"
    }
]
```


### Working with JSON
#### Parse Json Action
use parse json action to parse request in json

#### Working with object variable
Initalize with type being `Object`

##### Set / add property
Add a placeholder variable of type `Object`, then use the following function `addPropety(object, propName, propValue)` or `setPropety(object, propName, propValue)`

#### Convert JSON Stringify String into Flow Object
wraps it up in the stirng `JSON`, for example `JSON(body('Checkout_Parsed_Json')?['d']?['CanvasContent1'])`

#### Convert JSON Object to JSON string (aka JSON.stringify)
wraps it up in the `string(...)`

### How to construct profile image URL using the email

Given an email, this can be done via this URL format

```
 "/_layouts/15/userphoto.aspx?size=L&username=bswagger@sharepoint-test.com"
```


### API used to get the site user by email

`GET ${this.props.siteUrl}/_api/Web/SiteUsers?$filter=Email eq '${emailString}'`

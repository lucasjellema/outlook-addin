# The Image Inserter Outlook Addin

This article describes the creation of a simple Outlook Addin. It acts on a message that is being composed in the message editor. When the Addin is activated, it will cause an image - and a very pretty one! - to be inserted in the message under construction at the position of the cursor. It will provide a starting point for much more complex and interesting Addins later on.

The steps to go through in order to create, test-run and deploy the Addin are outline below.

## Prepare your environment

Make sure you have VS Code installed. Also ensure that you have Node set up: Node.js (the latest LTS version). Visit [the Node.js site](https://nodejs.org/) to download and install the right version for your operating system. And work on a machine that has the Outlook Client installed. (note: you can work with Outlook Web Client as well, with a flow that is a little bit less smooth).

Run VS Code. Open a Terminal (I prefer the Bash terminal).

Install the latest version of Yeoman and the Yeoman generator for Office Add-ins. To install these tools globally, run the following command via the command prompt.
```
npm install -g yo generator-officenpm install -g yo generator-office
```

This is the general preparation, one time only and the same of all types of Office Addin development (including Word, Excel, Powerpoint and OneNote).


## Generate the Skaffold for the ImageInserter Addin

Run the following command to create an add-in project using the Yeoman generator. A folder that contains the project will be added to the current directory.
```
yo office
```
When prompted, provide the following information to create your add-in project.
```
Choose a project type: Office Add-in Task Pane project
? Choose a script type: JavaScript
? What do you want to name your add-in? ImageInserter
? Which Office client application would you like to support? Outlook
? Which manifest type would you like to use? XML manifest
```

![](images/run-yo-office.png)

The directory ImageInserter is created, with artifacts that form the Addin.

![](images/artifacts-of-addin.png)

To make sure all modules are installed correctly, navigate into the new directory and run `npm install`:
```
cd ImageInserter
npm install
```
![](images/npom-install.png)

## Refine the generated default Addin

Inside VS Code, edit the file `manifest.xml`.

Update the elements ProviderName and Description with apprpriate values.

Change the ExtensionPoint element's type attribute to `MessageComposeCommandSurface` to make this addin active for new email composition.
```
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```

Add Rule element (type ItemIs)
```
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
```
inside the <Rule> of type RuleCollection, to make it read:
```
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>
```
Update the file  `src/taskpane.js`: 

Add function insertImage at the end of the file:
```
function insertImage() {
  const imageDataUrl = 'https://www.thewowstyle.com/wp-content/uploads/2015/01/images-of-nature-4.jpg'
 
  // Create an HTML image element
  const imgElement = `<img src="${imageDataUrl}" alt="Inserted Image" />`;

  // Insert the image HTML at the cursor position
  Office.context.mailbox.item.body.setSelectedDataAsync(
      imgElement,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Image inserted successfully.");
          } else {
              console.error("Failed to insert image: " + asyncResult.error.message);
          }
      }
  );
}
```

and change function `run` to:
```
export async function run() {
  /**
   * Insert your Outlook code here
   */

 insertImage();
}

```

## Test run the Addin

Now run `npm start`. Type `N` when asked *? Allow localhost loopback for Microsoft Edge WebView? No*

This will:
* run a web server that serves the files the addin is made up of
* add the addin to your local Outlook client, referring to localhost where the web server is providing the addin
* run Outlook - which now has the Addin enabled

![](images/test-run-addin.png)

Goto Outlook. Create a new Email. Start typing your message. 

When you want to insert an image: Click on the Apps icon. You will find the *ImageInserter*. 
![](images/new-message-addins.png)

Click on the ImageInserter. A dropdown menu pops up. 
![](images/image-inserrter-dropdown.png)
Click on Show Taskpane.

The taskpane (defined in taskpane.html) is shown on the right hand side of the screen:
![](images/image-inserter-taskpane.png)

Click on the *Run* link in the ImageInserter Taskpane.

The image defined in the function is now added in the email editor at the position of the cursor:
![](images/image-inserted.png)

If you now make any change in taskpane.html or taskpane.js, that change is reflected in your Outlook client immediately.

Change for example the <body> element, replacing it with this content:
```
<body class="ms-font-m ms-welcome ms-Fabric">
    <header class="ms-welcome__header ms-bgColor-neutralLighter">
        <img width="90" height="90" src="../../assets/logo-filled.png" alt="Contoso" title="Contoso" />
        <h1 class="ms-font-su">The Image Inserter</h1>
    </header>
    <section id="sideload-msg" class="ms-welcome__main">
        <h2 class="ms-font-xl">Please <a target="_blank" href="https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing">sideload</a> your add-in to see app body.</h2>
    </section>
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Brighten up your emails! </h2>
        <ul class="ms-List ms-welcome__features">
            <li class="ms-ListItem">
                <i class="ms-Icon ms-Icon--Ribbon ms-font-xl"></i>
                <span class="ms-font-m">Add wonderful images</span>
            </li>
            <li class="ms-ListItem">
                <i class="ms-Icon ms-Icon--Design ms-font-xl"></i>
                <span class="ms-font-m">Prepare stunning emails for friends and business partners</span>
            </li>
        </ul>
        <p class="ms-font-l">Click the button to inject an image into your email</p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Insert Image</span>
        </div>
        <p><label id="item-subject"></label></p>
    </main>
</body>
```
and check the Outlook client again:
![](images/refreshed-taskpane-html.png)


## Add Addin for real

The Addin is now running in a typical development setup. Only during the current Outlook client's session and only because the local webserver is running do you have access to the Addin. To make the Addin a real fixture in your Outlook client - across sessions - and also to make it available to others, the Addin must be served centrally from a web server (or more formally: it should be published on Microsoft's AppSource program or your own organization's Microsoft 365 admin center).

An easy way of making an Addin available across Outlook sessions and clients/users is by publishing it on GitHub Pages. Here are the steps for this:


# Resources

Build your first addin - https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=yeomangenerator

Tutorial: Build a message compose Outlook add-in - https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/outlook-tutorial?tabs=jsonmanifest


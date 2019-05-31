This is sample demo to get view count for files in a SharePoint document library.

The webpart depends on jQuery https://code.jquery.com/jquery-2.1.1.min.js and https://cdnjs.cloudflare.com/ajax/libs/jquery-treegrid/0.2.0/js/jquery.treegrid.min.js

After you configured the webpart library property, refresh the page to check the result.
Sample result:

https://1drv.ms/u/s!ArH29oxgtifigQ0-7kGOjPVV7Zte

If you have a large library, you need optimize the logic to paging/batch for better performance.

## Used SharePoint Framework Version
1.8.2

### Local testing

- Clone this repository
- In the command line run:
  - `npm install`
  - `gulp serve`

### Deploy
* gulp clean
* gulp bundle --ship
* gulp package-solution --ship
* Upload .sppkg file from sharepoint\solution to your tenant App Catalog
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog
* Add the web part to a site collection, and test it on a page

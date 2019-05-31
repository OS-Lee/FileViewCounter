This is sample demo to get view count for files in a SharePoint document library.

Used SharePoint Framework Version
1.8.2

###Local testing
Clone this repository
In the command line run:
npm install
gulp serve

###Deploy
gulp clean
gulp bundle --ship
gulp package-solution --ship
Upload .sppkg file from sharepoint\solution to your tenant App Catalog
E.g.: https://<tenant>.sharepoint.com/sites/AppCatalog/AppCatalog
Add the web part to a site collection, and test it on a page

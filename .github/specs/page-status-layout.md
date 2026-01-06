# Page Status PnP Modern Search Custom Layout

## Intent

* Allow PnP Modern Search solution users to display SharePoint pages and SharePoint news.
* Allow PnP Modern Search solution users to select a custom layout on the Search results component.
* Custom layout includes a custom web component to handle specific actions (delete, edit, details).

## Scope

* Custom layout and custom web component in SPFx to extend PnP Modern Search Solution.

## Must do

* Display all SharePoint pages and news posts match the query entered in the PnP Modern Search web part (search results)
* For each page, layout must display the thumbnail image and the title below the image.
* On each thumbnail image display an indicator (choose best option between ribbon or icon or suggest better option) to indicate if the page is in Draft or Scheduled to be published (nothing if page is already published)
* For each page (choose best location) use a custom web component to display a delete action icon, an edit action icon and a details action icon.
* If user clicks on Delete action icon, a pop-up (modal) open asking for confirmation to delete the page. If user clicks on "OK", the page is deleted.
* If user clicks on Edit action icon, the page opens in Edit Mode.
* If user cliks on Details action icon, a pop-up (modal) opens to display addtional page data (number of views, like count, comment count, last modified date). A loading screen must be there while data loads.
* Additional page data must only be loaded if user clicks on Details

## Must not do

* Must not load all retrieved pages additional data at once/

## Constraints

* Use PnPjs library when possible.
* For UI components, use fluentUI/react library
* Use .github/specs/layout.html as starting point for the layout


## Definition of done

* The code builds successfully.
* All global rules are respected.

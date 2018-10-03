# OneNote JavaScript API overview

Applies to: OneNote Online

The following links show the high level OneNote objects available in the API. Each object page link contains a description of the properties, events, and methods available on the object. Explore these links to learn more. 
	
- [Application](/javascript/api/onenote/onenote.application): The top-level object used to access all globally addressable OneNote objects, such as the active notebook and the active section.

- [Notebook](/javascript/api/onenote/onenote.notebook): A notebook. Notebooks contain section groups and sections.
    - [NotebookCollection](/javascript/api/onenote/onenote.notebookcollection): A collection of notebooks.

- [SectionGroup](/javascript/api/onenote/onenote.sectiongroup): A section group. Section groups contain section groups and sections.
    - [SectionGroupCollection](/javascript/api/onenote/onenote.sectiongroupcollection): A collection of section groups.

- [Section](/javascript/api/onenote/onenote.section): A section. Sections contain pages.
    - [SectionCollection](/javascript/api/onenote/onenote.sectioncollection): A collection of sections.

- [Page](/javascript/api/onenote/onenote.page): A page. Pages contain PageContent objects.
    - [PageCollection](/javascript/api/onenote/onenote.pagecollection): A collection of pages.

- [PageContent](/javascript/api/onenote/onenote.pagecontent): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.
    - [PageContentCollection](/javascript/api/onenote/onenote.pagecontentcollection): A collection of PageContent objects, which represents the contents of a page.

- [Outline](/javascript/api/onenote/onenote.outline): A container for Paragraph objects. An Outline is a direct child of a PageContent object.

- [Image](/javascript/api/onenote/onenote.image): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.

- [Paragraph](/javascript/api/onenote/onenote.paragraph): A container for the visible content on a page. A Paragraph is a direct child of an Outline.
    - [ParagraphCollection](/javascript/api/onenote/onenote.paragraphcollection): A collection of Paragraph objects in an Outline.

- [RichText](/javascript/api/onenote/onenote.richtext): A RichText object.

- [Table](/javascript/api/onenote/onenote.table): A container for TableRow objects.

- [TableRow](/javascript/api/onenote/onenote.tablerow): A container for TableCell objects.
    - [TableRowCollection](/javascript/api/onenote/onenote.tablerowcollection): A collection of TableRow objects in a Table.
 
- [TableCell](/javascript/api/onenote/onenote.tablecell): A container for Paragraph objects.
    - [TableCellCollection](/javascript/api/onenote/onenote.tablecellcollection): A collection of TableCell objects in a TableRow.

## OneNote JavaScript API reference

For detailed information about OneNote JavaScript API, see the [OneNote JavaScript API reference documentation](/javascript/api/onenote).

## See also

- [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview)
- [Build your first OneNote add-in](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-getting-started)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)

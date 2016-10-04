# OneNote add-ins JavaScript API reference

*Applies to: OneNote Online*

The links below show the high level OneNote objects available in the API. Each object page link contains a description of the properties, relationships, and methods available on the object. Explore the links below to learn more. 
	
- [Application](application.md): The top-level object used to access all globally addressable OneNote objects, such as the active notebook, and the active section.

- [Notebook](notebook.md): A notebook. Notebooks contain section groups and sections.

   - [NotebookCollection](notebookcollection.md): A collection of notebooks.

- [SectionGroup](sectiongroup.md): A section group. Section groups contain section groups and sections.

   - [SectionGroupCollection](sectiongroupcollection.md): A collection of section groups.

- [Section](section.md): A section. Sections contain pages.

   - [SectionCollection](sectioncollection.md): A collection of sections.

- [Page](page.md): A page. Pages contain PageContent objects.

   - [PageCollection](pagecollection.md): A collection of pages.

- [PageContent](pagecontent.md): A top-level region on a page that contains content types such as Outline or Image. A PageContent object can be assigned a position on the page.

   - [PageContentCollection](pagecontentcollection.md): A collection of PageContent objects, which represents the contents of a page.

- [Outline](outline.md): A container for Paragraph objects. An Outline is a direct child of a PageContent object.

- [Image](image.md): An Image object. An Image can be a direct child of a PageContent object or a Paragraph.

- [Paragraph](paragraph.md): A container for the visible content on a page. A Paragraph is a direct child of an Outline.

  - [ParagraphCollection](paragraphcollection.md): A collection of Paragraph objects in an Outline.

- [RichText](richtext.md): A RichText object.

- [Table](table.md): A container for TableRow objects.

- [TableRow](tablerow.md): A container for TableCell objects.

  - [TableRowCollection](tablerowcollection.md): A collection of TableRow objects in a Table.
 
- [TableCell](tablecell.md): A container for Paragraph objects.

  - [TableCellCollection](tablecellcollection.md): A collection of TableCell objects in a TableRow.
		
## Additional resources

- [OneNote JavaScript API programming overview](../../docs/onenote/onenote-add-ins-programming-overview.md)
- [Build your first OneNote add-in](../../docs/onenote/onenote-add-ins-getting-started.md)
- [Rubric Grader sample](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office Add-ins platform overview](https://dev.office.com/docs/add-ins/overview/office-add-ins)

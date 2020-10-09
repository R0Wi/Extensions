# Leviasan.Extension.EPPlus

### How use it:
1. Determine the data model that will be placed on the Excel sheet.
2. In the template, in the cells create keys in the format [name of the model property].
3. Use the extension methods of the ExcelWorksheet object: LoadFromDictionary or LoadFromCollection.

Using attributes over model properties will allow to control data in the range:
- [NumberformatAttribute] - data format.
- [RangeStyleAttribute] - styles.
- [HyperlinkAttribute] - link behavior.
- [DisplayNameAttribute] - the name of the column in the table when it is filled using the LoadFromCollection method with the "printHeaders = true" parameter.
- [GroupingAttribute] - data grouping.
- [DescriptionAttribute] - comment.
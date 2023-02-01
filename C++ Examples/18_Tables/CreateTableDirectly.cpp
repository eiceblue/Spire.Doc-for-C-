#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;


int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateTableDirectly.docx";

	//Create a Word document
	Document* doc = new Document();

	//Add a section
	Section* section = doc->AddSection();

	//Create a table 
	Table* table = new Table(doc);
	//Set the width of table
	table->SetPreferredWidth(new PreferredWidth(WidthType::Percentage, 100));
	//Set the border of table
	table->GetTableFormat()->GetBorders()->SetBorderType(BorderStyle::Single);

	//Create a table row
	TableRow* row = new TableRow(doc);
	row->SetHeight(50.0f);
	table->GetRows()->Add(row);

	//Create a table cell
	TableCell* cell1 = new TableCell(doc);
	//Add a paragraph
	Paragraph* para1 = cell1->AddParagraph();
	//Append text in the paragraph
	para1->AppendText(L"Row 1, Cell 1");
	//Set the horizontal alignment of paragrah
	para1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	//Set the background color of cell
	cell1->GetCellFormat()->SetBackColor(Color::GetCadetBlue());
	//Set the vertical alignment of paragraph
	cell1->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	row->GetCells()->Add(cell1);

	//Create a table cell
	TableCell* cell2 = new TableCell(doc);
	Paragraph* para2 = cell2->AddParagraph();
	para2->AppendText(L"Row 1, Cell 2");
	para2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	cell2->GetCellFormat()->SetBackColor(Color::GetCadetBlue());
	cell2->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	row->GetCells()->Add(cell2);

	//Add the table in the section
	section->GetTables()->Add(table);

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}

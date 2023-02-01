#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;


int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RepeatRowOnEachPage.docx";

	//Create word document
	Document* document = new Document();

	//Create a new section
	Section* section = document->AddSection();

	//Create a table width default borders
	Table* table = section->AddTable(true);
	//Set table with to 100%
	PreferredWidth* width = new PreferredWidth(WidthType::Percentage, 100);
	table->SetPreferredWidth(width);

	//Add a new row 
	TableRow* row = table->AddRow();
	//Set the row as a table header 
	row->SetIsHeader(width);
	//Set the backcolor of row
	row->GetRowFormat()->SetBackColor(Color::GetLightGray());
	//Add a new cell for row
	TableCell* cell = row->AddCell();
	cell->SetCellWidth(100, CellWidthType::Percentage);
	//Add a paragraph for cell to put some data
	Paragraph* parapraph = cell->AddParagraph();
	//Add text 
	parapraph->AppendText(L"Row Header 1");
	//Set paragraph horizontal center alignment
	parapraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	row = table->AddRow(false, 1);
	row->SetIsHeader(true);
	row->GetRowFormat()->SetBackColor(Color::GetIvory());
	//Set row height
	row->SetHeight(30);
	cell = row->GetCells()->GetItem(0);
	cell->SetCellWidth(100, CellWidthType::Percentage);
	//Set cell vertical middle alignment
	cell->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	//Add a paragraph for cell to put some data
	parapraph = cell->AddParagraph();
	//Add text 
	parapraph->AppendText(L"Row Header 2");
	parapraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Add many common rows 
	for (int i = 0; i < 70; i++)
	{
		row = table->AddRow(false, 2);
		cell = row->GetCells()->GetItem(0);
		//Set cell width to 50% of table width
		cell->SetCellWidth(50, CellWidthType::Percentage);
		cell->AddParagraph()->AppendText(L"Column 1 Text");
		cell = row->GetCells()->GetItem(1);
		cell->SetCellWidth(50, CellWidthType::Percentage);
		cell->AddParagraph()->AppendText(L"Column 2 Text");
	}
	//Set cell backcolor
	for (int j = 1; j < table->GetRows()->GetCount(); j++)
	{
		if (j % 2 == 0)
		{
			TableRow* row2 = table->GetRows()->GetItem(j);
			for (int f = 0; f < row2->GetCells()->GetCount(); f++)
			{
				row2->GetCells()->GetItem(f)->GetCellFormat()->SetBackColor(Color::GetLightBlue());
			}
		}

	}

	//Save file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

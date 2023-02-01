#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DifferentBorders.docx";

	//Open a Word document as template
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	Table* table = dynamic_cast<Table*>(document->GetSections()->GetItem(0)->GetTables()->GetItemInTableCollection(0));

	//Set borders of table
	setTableBorders(table);

	//Set borders of cell
	setCellBorders(table->GetRows()->GetItem(2)->GetCells()->GetItem(0));

	//Save and launch document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void setTableBorders(Table* table)
{
	table->GetTableFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetTableFormat()->GetBorders()->SetLineWidth(3.0F);
	table->GetTableFormat()->GetBorders()->SetColor(Color::GetRed());
}

void setCellBorders(TableCell* tableCell)
{
	tableCell->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::DotDash);
	tableCell->GetCellFormat()->GetBorders()->SetLineWidth(1.0F);
	tableCell->GetCellFormat()->GetBorders()->SetColor(Color::GetGreen());
}
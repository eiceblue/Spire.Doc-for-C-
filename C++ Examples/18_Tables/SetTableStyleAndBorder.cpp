#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTableStyleAndBorder.docx";

	//Create a document and load file
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	Section* section = document->GetSections()->GetItem(0);

	//Get the first table
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	//Apply the table style
	table->ApplyStyle(DefaultTableStyle::ColorfulList);

	//Set right border of table
	table->GetTableFormat()->GetBorders()->GetRight()->SetBorderType(BorderStyle::Hairline);
	table->GetTableFormat()->GetBorders()->GetRight()->SetLineWidth(1.0F);
	table->GetTableFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());

	//Set top border of table
	table->GetTableFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Hairline);
	table->GetTableFormat()->GetBorders()->GetTop()->SetLineWidth(1.0F);
	table->GetTableFormat()->GetBorders()->GetTop()->SetColor(Color::GetGreen());

	//Set left border of table
	table->GetTableFormat()->GetBorders()->GetLeft()->SetBorderType(BorderStyle::Hairline);
	table->GetTableFormat()->GetBorders()->GetLeft()->SetLineWidth(1.0F);
	table->GetTableFormat()->GetBorders()->GetLeft()->SetColor(Color::GetYellow());

	//Set bottom border is none
	table->GetTableFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::DotDash);

	//Set vertical and horizontal border 
	table->GetTableFormat()->GetBorders()->GetVertical()->SetBorderType(BorderStyle::Dot);
	table->GetTableFormat()->GetBorders()->GetHorizontal()->SetBorderType(BorderStyle::None);
	table->GetTableFormat()->GetBorders()->GetVertical()->SetColor(Color::GetOrange());

	//Save the file and launch it
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

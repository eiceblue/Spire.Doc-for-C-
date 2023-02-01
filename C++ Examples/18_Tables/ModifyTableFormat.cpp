#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ModifyTableFormat.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ModifyTableFormat.docx";

	//Load Word document from disk
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = document->GetSections()->GetItem(0);

	//Get tables 
	Table* tb1 = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));
	Table* tb2 = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(1));
	Table* tb3 = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(2));

	MoidyTableFormat(tb1);
	ModifyRowFormat(tb2);
	ModifyCellFormat(tb3);

	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

void MoidyTableFormat(Table* table)
{
	//Set table width
	table->SetPreferredWidth(new PreferredWidth(WidthType::Twip, static_cast<short>(6000)));

	//Apply style for table
	table->ApplyStyle(DefaultTableStyle::ColorfulGridAccent3);

	//Set table padding
	table->GetTableFormat()->GetPaddings()->SetAll(5);

	//Set table title and description
	table->SetTitle(L"Spire.Doc for C++");
	table->SetTableDescription(L"Spire.Doc for C++ is a professional Word C++ library");
}

void ModifyRowFormat(Table* table)
{
	//Set cell spacing
	table->GetRows()->GetItem(0)->GetRowFormat()->SetCellSpacing(2);

	//Set row height
	table->GetRows()->GetItem(1)->SetHeightType(TableRowHeightType::Exactly);
	table->GetRows()->GetItem(1)->SetHeight(20.0f);

	//Set background color
	table->GetRows()->GetItem(2)->GetRowFormat()->SetBackColor(Color::GetDarkSeaGreen());
}

void ModifyCellFormat(Table* table)
{
	//Set alignment
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetParagraphs()->GetItem(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Set background color
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->GetCellFormat()->SetBackColor(Color::GetDarkSeaGreen());

	//Set cell border
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->SetBorderType(BorderStyle::Single);
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->SetLineWidth(1.0f);
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->GetLeft()->SetColor(Color::GetRed());
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->GetRight()->SetColor(Color::GetRed());
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->GetTop()->SetColor(Color::GetRed());
	table->GetRows()->GetItem(2)->GetCells()->GetItem(0)->GetCellFormat()->GetBorders()->GetBottom()->SetColor(Color::GetRed());

	//Set text direction
	table->GetRows()->GetItem(3)->GetCells()->GetItem(0)->GetCellFormat()->SetTextDirection(TextDirection::RightToLeft);
}
#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetVerticalAlignment.docx";

	//Create a new Word document and add a new section
	Document* doc = new Document();
	Section* section = doc->AddSection();

	//Add a table with 3 columns and 3 rows
	Table* table = section->AddTable(true);
	table->ResetCells(3, 3);

	//Merge rows
	table->ApplyVerticalMerge(0, 0, 2);

	//Set the vertical alignment for each cell, default is top
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(0)->GetCells()->GetItem(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
	table->GetRows()->GetItem(0)->GetCells()->GetItem(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Top);
	table->GetRows()->GetItem(1)->GetCells()->GetItem(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(1)->GetCells()->GetItem(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(2)->GetCells()->GetItem(1)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);
	table->GetRows()->GetItem(2)->GetCells()->GetItem(2)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Bottom);

	//Inset a picture into the table cell
	Paragraph* paraPic = table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->AddParagraph();
	DocPicture* pic = paraPic->AppendPicture(Image::FromFile(inputFile.c_str()));
	//Create data
	vector<vector<wstring>> data =
	{
		{L"", L"Spire.Office", L"Spire.DataExport"},
		{L"", L"Spire.Doc", L"Spire.DocViewer"},
		{L"", L"Spire.XLS", L"Spire.PDF"}
	};

	//Append data to table
	for (int r = 0; r < 3; r++)
	{
		TableRow* dataRow = table->GetRows()->GetItem(r);
		dataRow->SetHeight(50);
		for (int c = 0; c < 3; c++)
		{
			if (c == 1)
			{
				Paragraph* par = dataRow->GetCells()->GetItem(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				dataRow->GetCells()->GetItem(c)->SetWidth((section->GetPageSetup()->GetClientWidth()) / 2);
			}
			if (c == 2)
			{
				Paragraph* par = dataRow->GetCells()->GetItem(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				dataRow->GetCells()->GetItem(c)->SetWidth((section->GetPageSetup()->GetClientWidth()) / 2);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

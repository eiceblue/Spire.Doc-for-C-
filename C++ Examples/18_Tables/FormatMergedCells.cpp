#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormatMergedCells.docx";

	//Create word document
	Document* document = new Document();

	//Add a new section
	Section* section = document->AddSection();

	//Add a new table
	Table* table = AddTable(section);

	//Create a new style
	ParagraphStyle* style = new ParagraphStyle(document);
	style->SetName(L"Style");
	style->GetCharacterFormat()->SetTextColor(Color::GetDeepSkyBlue());
	style->GetCharacterFormat()->SetItalic(true);
	style->GetCharacterFormat()->SetBold(true);
	style->GetCharacterFormat()->SetFontSize(13);
	document->GetStyles()->Add(style);

	//Merge cell horizontally
	table->ApplyHorizontalMerge(0, 0, 1);
	//Apply style
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetParagraphs()->GetItem(0)->ApplyStyle(style->GetName());
	//Set vertical and horizontal alignment
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->GetParagraphs()->GetItem(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Merge cell vertically
	table->ApplyVerticalMerge(0, 1, 3);
	//Apply style
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->GetParagraphs()->GetItem(0)->ApplyStyle(style->GetName());
	//Set vertical and horizontal alignment
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->GetParagraphs()->GetItem(0)->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
	//Set column width
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->SetCellWidth(20, CellWidthType::Percentage);
	//Save and launch document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

Table* AddTable(Section* section)
{
	Table* table = section->AddTable(true);
	table->ResetCells(4, 3);
	//Table data
	vector<vector<wstring>> data =
	{
		{L"Product", L"", L"Price"},
		{L"Spire.Doc", L"Pro Edition", L"$799"},
		{L"", L"Standard Edition", L"$599"},
		{L"", L"Free Edition", L"$0"}
	};
	for (int r = 0; r < data.size(); r++)
	{
		TableRow* dataRow = table->GetRows()->GetItem(r);
		dataRow->SetHeight(20);
		dataRow->SetHeightType(TableRowHeightType::Exactly);
		dataRow->GetRowFormat()->SetBackColor(Color::Empty());
		for (int c = 0; c < data[r].size(); c++)
		{
			if (!data[r][c].empty()) {
				TextRange* range = dataRow->GetCells()->GetItem(c)->AddParagraph()->AppendText(data[r][c].c_str());
				range->GetCharacterFormat()->SetFontName(L"Arial");
			}
		}
	}
	return table;
}

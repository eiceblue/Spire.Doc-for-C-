#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Word.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetOutsidePosition.docx";

	//Create a new word document and add new section
	Document* doc = new Document();
	Section* sec = doc->AddSection();

	//Get header
	HeaderFooter* header = doc->GetSections()->GetItem(0)->GetHeadersFooters()->GetHeader();

	//Add new paragraph on header and set HorizontalAlignment of the paragraph as left
	Paragraph* paragraph = header->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	//Load an image for the paragraph
	DocPicture* headerimage = paragraph->AppendPicture(Image::FromFile(inputFile.c_str()));
	//Add a table of 4 rows and 2 columns
	Table* table = header->AddTable();
	table->ResetCells(4, 2);

	//Set the position of the table to the right of the image
	table->GetTableFormat()->SetWrapTextAround(true);
	table->GetTableFormat()->GetPositioning()->SetHorizPositionAbs(HorizontalPosition::Outside);
	table->GetTableFormat()->GetPositioning()->SetVertRelationTo(VerticalRelation::Margin);
	table->GetTableFormat()->GetPositioning()->SetVertPosition(43);

	//Add contents for the table
	vector<vector<wstring>> data =
	{
		{L"Spire.Doc.left", L"Spire XLS.right"},
		{L"Spire.Presentatio.left", L"Spire.PDF.right"},
		{L"Spire.DataExport.left", L"Spire.PDFViewe.right"},
		{L"Spire.DocViewer.left", L"Spire.BarCode.right"}
	};

	for (int r = 0; r < 4; r++)
	{
		TableRow* dataRow = table->GetRows()->GetItem(r);
		for (int c = 0; c < 2; c++)
		{
			if (c == 0)
			{
				Paragraph* par = dataRow->GetCells()->GetItem(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
				dataRow->GetCells()->GetItem(c)->SetWidth(180);
			}
			else
			{
				Paragraph* par = dataRow->GetCells()->GetItem(c)->AddParagraph();
				par->AppendText(data[r][c].c_str());
				par->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
				dataRow->GetCells()->GetItem(c)->SetWidth(180);
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

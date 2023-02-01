#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetTextDirection.docx";

	//Create a new document
	Document* doc = new Document();

	//Add the first section
	Section* section1 = doc->AddSection();
	//Set text direction for all text in a section
	section1->SetTextDirection(TextDirection::RightToLeft);

	//Set Font Style and Size
	ParagraphStyle* style = new ParagraphStyle(doc);
	style->SetName(L"FontStyle");
	style->GetCharacterFormat()->SetFontName(L"Arial");
	style->GetCharacterFormat()->SetFontSize(15);
	doc->GetStyles()->Add(style);

	//Add two paragraphs and apply the font style
	Paragraph* p = section1->AddParagraph();
	p->AppendText(L"Only Spire.Doc, no Microsoft Office automation");
	p->ApplyStyle(style->GetName());
	p = section1->AddParagraph();
	p->AppendText(L"Convert file documents with high quality");
	p->ApplyStyle(style->GetName());

	//Set text direction for a part of text
	//Add the second section
	Section* section2 = doc->AddSection();
	//Add a table
	Table* table = section2->AddTable();
	table->ResetCells(1, 1);
	TableCell* cell = table->GetRows()->GetItem(0)->GetCells()->GetItem(0);
	table->GetRows()->GetItem(0)->SetHeight(150);
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->SetCellWidth(10, CellWidthType::Point);
	//Set vertical text direction of table
	cell->GetCellFormat()->SetTextDirection(TextDirection::RightToLeftRotated);
	cell->AddParagraph()->AppendText(L"This is vertical style");
	//Add a paragraph and set horizontal text direction
	p = section2->AddParagraph();
	p->AppendText(L"This is horizontal style");
	p->ApplyStyle(style->GetName());

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

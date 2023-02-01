#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"InsertEndnote.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertEndnote.docx";

	//Create a document and load file
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	Section* s = doc->GetSections()->GetItem(0);
	Paragraph* p = s->GetParagraphs()->GetItem(1);

	//add endnote
	Footnote* endnote = p->AppendFootnote(FootnoteType::Endnote);

	//append text
	TextRange* text = endnote->GetTextBody()->AddParagraph()->AppendText(L"Reference: Wikipedia");

	//set text format
	text->GetCharacterFormat()->SetFontName(L"Impact");
	text->GetCharacterFormat()->SetFontSize(14);
	text->GetCharacterFormat()->SetTextColor(Color::GetDarkOrange());

	//Set marker format of endnote
	endnote->GetMarkerCharacterFormat()->SetFontName(L"Calibri");
	endnote->GetMarkerCharacterFormat()->SetFontSize(25);
	endnote->GetMarkerCharacterFormat()->SetTextColor(Color::GetDarkBlue());

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

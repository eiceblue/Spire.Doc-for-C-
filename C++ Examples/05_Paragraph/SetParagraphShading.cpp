#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetParagraphShading.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());
	//Get a paragraph.
	Paragraph* paragaph = document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0);

	//Set background color for the paragraph.
	paragaph->GetFormat()->SetBackColor(Color::GetYellow());

	//Set background color for the selected text of paragraph.
	paragaph = document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(2);
	TextSelection* selection = paragaph->Find(L"Christmas", true, false);
	TextRange* range = selection->GetAsOneRange();
	range->GetCharacterFormat()->SetTextBackgroundColor(Color::GetYellow());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

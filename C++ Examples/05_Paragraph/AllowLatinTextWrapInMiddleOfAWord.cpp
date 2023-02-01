#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"AllowLatinTextWrapInMiddleOfAWord.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AllowLatinTextWrapInMiddleOfAWord.docx";

	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());
	Paragraph* para = document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0);
	//Allow Latin text to wrap in the middle of a word
	para->GetFormat()->SetWordWrap(false);
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceWithText.docx";

	//Create word document
	Document* document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Replace text
	document->Replace(L"word", L"ReplacedText", false, true);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
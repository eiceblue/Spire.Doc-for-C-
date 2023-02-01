#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractText.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetText.txt";

	//Load the document from disk.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//get text from document
	wstring text = document->GetText();

	//create a new TXT File to save extracted text
	wofstream write(outputFile);
	write << text;
	write.close();
	document->Close();
	delete document;
}
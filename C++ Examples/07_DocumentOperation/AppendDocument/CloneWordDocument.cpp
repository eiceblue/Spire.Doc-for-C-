#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneWordDocument.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Clone the word document.
	Document* newDoc = document->Clone();

	//Save the file.
	newDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	newDoc->Close();
	document->Close();
	delete document;
}
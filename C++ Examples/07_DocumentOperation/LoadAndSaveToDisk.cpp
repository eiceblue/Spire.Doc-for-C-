#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LoadAndSaveToDisk.docx";

	//Create a new document
	Document* doc = new Document();
	// Load the document from the absolute/relative path on disk.
	doc->LoadFromFile(inputFile.c_str());

	// Save the document to disk
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"IfFieldSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"UpdateFields.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Update fields
	document->SetIsUpdateFields(true);

	//Save doc file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
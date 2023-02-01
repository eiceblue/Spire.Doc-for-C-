#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TemplateWithPassword.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Decrypt.docx";

	//Create word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str(), FileFormat::Docx, L"E-iceblue");

	//Save as doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}


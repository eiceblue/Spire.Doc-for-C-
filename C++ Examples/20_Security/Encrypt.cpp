#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Encrypt.docx";

	//Create word document
	Document* document = new Document();

	//Load Word document.
	document->LoadFromFile(inputFile.c_str());

	//encrypt document with password specified by textBox1
	document->Encrypt(L"E-iceblue");

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

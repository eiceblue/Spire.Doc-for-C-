#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SpecifiedProtectionType.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Protect the Word file.
	document->Protect(ProtectionType::AllowOnlyReading, L"123456");

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

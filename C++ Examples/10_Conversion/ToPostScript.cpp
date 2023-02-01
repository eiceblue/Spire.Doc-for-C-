#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToPostScript.ps";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	doc->SaveToFile(outputFile.c_str(), FileFormat::PostScript);
	doc->Close();
	delete doc;
}

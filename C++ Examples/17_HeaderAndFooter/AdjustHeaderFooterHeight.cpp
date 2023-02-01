#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"HeaderAndFooter.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AdjustHeaderFooterHeight.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Adjust the height of headers in the section
	section->GetPageSetup()->SetHeaderDistance(100);

	//Adjust the height of footers in the section
	section->GetPageSetup()->SetFooterDistance(100);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

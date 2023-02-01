#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"DifferentPageSetup.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DifferentPageSetup.docx";


	//Open a Word document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the second section 
	Section* SectionTwo = doc->GetSections()->GetItem(1);

	//Set the orientation
	SectionTwo->GetPageSetup()->SetOrientation(PageOrientation::Landscape);

	doc->SaveToFile(outputFile.c_str());
	doc->Close();
	delete doc;
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SectionTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CloneSection.docx";

	//Load source file
	Document* srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile.c_str());

	//Create destination file
	Document* desDoc = new Document();

	Section* cloneSection = nullptr;
	for (int i = 0; i < srcDoc->GetSections()->GetCount(); i++)
	{
		Section* section = srcDoc->GetSections()->GetItem(i);
		//Clone section
		cloneSection = section->Clone();
		//Add the cloneSection in destination file
		desDoc->GetSections()->Add(cloneSection);
	}
	//Save the Word
	desDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	srcDoc->Close();
	desDoc->Close();
	delete srcDoc;
	delete desDoc;
}

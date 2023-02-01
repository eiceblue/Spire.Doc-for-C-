#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SectionTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddAndDeleteSections.docx";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	AddSection(doc);
	DeleteSection(doc);

	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();

	delete doc;
}

void AddSection(Document* doc)
{
	//Add a section
	doc->AddSection();
}

void DeleteSection(Document* doc)
{
	//Delete the last section
	doc->GetSections()->RemoveAt(doc->GetSections()->GetCount() - 1);
}

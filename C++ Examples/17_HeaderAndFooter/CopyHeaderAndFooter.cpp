#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"HeaderAndFooter.docx";
	wstring inputFile_1 = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyHeaderAndFooter.docx";

	//Load the source file
	Document* doc1 = new Document();
	doc1->LoadFromFile(inputFile.c_str());

	//Get the header section from the source document
	HeaderFooter* header = doc1->GetSections()->GetItem(0)->GetHeadersFooters()->GetHeader();

	//Load the destination file
	Document* doc2 = new Document();
	doc2->LoadFromFile(inputFile_1.c_str());

	//Copy each object in the header of source file to destination file
	for (int i = 0; i < doc2->GetSections()->GetCount(); i++)
	{
		Section* section = doc2->GetSections()->GetItem(i);
		for (int j = 0; j < header->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* obj = header->GetChildObjects()->GetItem(j);
			section->GetHeadersFooters()->GetHeader()->GetChildObjects()->Add(obj->Clone());
		}
	}

	//Save and launch document
	doc2->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc1->Close();
	doc2->Close();
	delete doc1;
	delete doc2;
}

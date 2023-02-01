#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Template_N2.docx";
	wstring inputFile_2 = input_path + L"Template_N3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"KeepSameFormat.docx";

	//Load the source document from disk
	Document* srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile_1.c_str());

	//Load the destination document from disk
	Document* destDoc = new Document();
	destDoc->LoadFromFile(inputFile_2.c_str());

	//Keep same format of source document
	srcDoc->SetKeepSameFormat(true);

	//Copy the sections of source document to destination document
	int sectionCount = srcDoc->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		Section* section = srcDoc->GetSections()->GetItem(i);
		destDoc->GetSections()->Add(section->Clone());
	}

	//Save the Word document
	destDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	srcDoc->Close();
	destDoc->Close();
	delete srcDoc;
	delete destDoc;
}
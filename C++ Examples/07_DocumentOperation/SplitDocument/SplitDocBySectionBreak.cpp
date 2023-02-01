#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_4.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SplitDocBySectionBreak/";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Define another new word document object.
	Document* newWord;

	//Split a Word document into multiple documents by section break.
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		wstring result = outputFile.c_str();
		result += L"SplitDocBySectionBreak_" + to_wstring(i) + L".docx";
		newWord = new Document();
		newWord->GetSections()->Add(document->GetSections()->GetItem(i)->Clone());

		//Save to file.
		newWord->SaveToFile(result.c_str());
		newWord->Close();
		delete newWord;
	}
}
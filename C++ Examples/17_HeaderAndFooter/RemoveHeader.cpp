#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"HeaderAndFooter.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveHeader.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section of the document
	Section* section = doc->GetSections()->GetItem(0);

	//Traverse the word document and clear all headers in different type
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		Paragraph* para = section->GetParagraphs()->GetItem(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* obj = para->GetChildObjects()->GetItem(j);
			//Clear header in the first page
			HeaderFooter* header;
			header = section->GetHeadersFooters()->GetFirstPageHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
			//Clear header in the odd page
			header = section->GetHeadersFooters()->GetOddHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
			//Clear header in the even page
			header = section->GetHeadersFooters()->GetEvenHeader();
			if (header != nullptr)
			{
				header->GetChildObjects()->Clear();
			}
		}
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

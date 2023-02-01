#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"HeaderAndFooter.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveFooter.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = doc->GetSections()->GetItem(0);

	//Traverse the word document and clear all footers in different type
	for (int i = 0; i < section->GetParagraphs()->GetCount(); i++)
	{
		Paragraph* para = section->GetParagraphs()->GetItem(i);
		for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
		{
			DocumentObject* obj = para->GetChildObjects()->GetItem(j);
			//Clear footer in the first page
			HeaderFooter* footer;
			footer = section->GetHeadersFooters()->GetFirstPageFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
			//Clear footer in the odd page
			footer = section->GetHeadersFooters()->GetOddFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
			//Clear footer in the even page
			footer = section->GetHeadersFooters()->GetEvenFooter();
			if (footer != nullptr)
			{
				footer->GetChildObjects()->Clear();
			}
		}
	}


	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}

#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Styles.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RetrieveStyle.txt";
	
	//Load a template document 
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Traverse all paragraphs in the document and get their style names through StyleName property
	wstring styleName = L"";
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* section = doc->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* paragraph = section->GetParagraphs()->GetItem(j);
			styleName.append(paragraph->GetStyleName());
			styleName.append(L"\n");
		}
	}

	//Save and launch document
	wofstream foo(outputFile);
	foo << styleName;
	foo.close();
	doc->Close();
	delete doc;
}
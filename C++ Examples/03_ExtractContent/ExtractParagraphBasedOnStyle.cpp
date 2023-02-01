#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ExtractParagraphBasedOnStyle.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractParagraphBasedOnStyle.txt";

	//Create a new document
	Document* document = new Document();
	wstring styleName1 = L"Heading1";
	wstring* style1Text = new wstring();
	//Load file from disk
	document->LoadFromFile(inputFile.c_str());
	style1Text->append(L"The following is the content of the paragraph with the style name " + styleName1 + L": \r\n");
	//Extrct paragraph based on style
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		Section* section = document->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* paragraph = section->GetParagraphs()->GetItem(j);
			if (paragraph->GetStyleName() != nullptr && paragraph->GetStyleName() == styleName1)
			{
				style1Text->append(paragraph->GetText());
			}
		}
	}

	wofstream write(outputFile);
	write << style1Text->c_str();
	write.close();
	document->Close();
	delete document;
	delete style1Text;
}

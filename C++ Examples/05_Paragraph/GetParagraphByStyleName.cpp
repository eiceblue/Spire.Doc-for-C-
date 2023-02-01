#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetParagraphByStyleName.txt";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	wstring* content = new wstring();
	content->append(L"Get paragraphs by style name \"Heading1\": ");
	content->append(L"\n");

	//Get paragraphs by style name.
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		Section* section = document->GetSections()->GetItem(i);
		for (int j = 0; j < section->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* paragraph = section->GetParagraphs()->GetItem(j);
			wstring style_name = paragraph->GetStyleName();
			if (style_name.compare(L"Heading1") == 0)
			{
				content->append(paragraph->GetText());
				content->append(L"\n");
			}
		}
	}

	//Save to file.
	wofstream write(outputFile.c_str());
	write << content->c_str();
	write.close();
	document->Close();
	delete document;
	delete content;
}

#include "pch.h"
#include <regex>

using namespace Spire::Doc;
int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableOfContent.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveTableOfContent.docx";

	//Create a document
	Document* document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first GetBody() from the first section
	Body* body = document->GetSections()->GetItem(0)->GetBody();

	//Remove TOC from first GetBody()
	wregex regexStr(L"TOC\\w+");
	for (int i = 0; i < body->GetParagraphs()->GetCount(); i++)
	{
		wstring styleName = body->GetParagraphs()->GetItem(i)->GetStyleName();

		if (regex_match(styleName.c_str(), regexStr))
		{
			body->GetParagraphs()->RemoveAt(i);
			i--;
		}
	}

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}


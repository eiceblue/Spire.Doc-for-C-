#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"BookmarkTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractBookmarkText.txt";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Creates a BookmarkNavigator instance to access the bookmark
	BookmarksNavigator* navigator = new BookmarksNavigator(doc);

	//Locate a specific bookmark by bookmark name
	navigator->MoveToBookmark(L"Content");
	TextBodyPart* textBodyPart = navigator->GetBookmarkContent();

	//Iterate through the items in the bookmark content to get the text
	wstring text = L"";
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		auto item = textBodyPart->GetBodyItems()->GetItem(i);
		if (dynamic_cast<Paragraph*>(item) != nullptr)
		{
			Paragraph* paragraph = (dynamic_cast<Paragraph*>(item));
			for (int j = 0; j < paragraph->GetChildObjects()->GetCount(); j++)
			{
				DocumentObject* childObject = paragraph->GetChildObjects()->GetItem(j);
				if (dynamic_cast<TextRange*>(childObject) != nullptr)
				{
					text += (dynamic_cast<TextRange*>(childObject))->GetText();
				}
			}
		}
	}

	//Save to TXT File and launch it
	wofstream foo(outputFile);
	foo << text;
	foo.close();
	doc->Close();
	delete doc;
	delete navigator;
}

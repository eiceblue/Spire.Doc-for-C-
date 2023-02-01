#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring insertFile = input_path + L"Insert.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertDocAtBookmark.docx";

	//Create the first document
	Document* document1 = new Document();

	//Load the first document from disk.
	document1->LoadFromFile(inputFile.c_str());

	//Create the second document
	Document* document2 = new Document();

	//Load the second document from disk.
	document2->LoadFromFile(insertFile.c_str());

	//Get the first section of the first document 
	Section* section1 = document1->GetSections()->GetItem(0);

	//Locate the bookmark
	BookmarksNavigator* bn = new BookmarksNavigator(document1);

	//Find bookmark by name
	bn->MoveToBookmark(L"Test", true, true);

	//Get bookmarkStart
	BookmarkStart* start = bn->GetCurrentBookmark()->GetBookmarkStart();

	//Get the owner paragraph
	Paragraph* para = start->GetOwnerParagraph();

	//Get the para index
	int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

	//Insert the paragraphs of document2
	for (int i = 0; i < document2->GetSections()->GetCount(); i++)
	{
		Section* section2 = document2->GetSections()->GetItem(i);
		for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* paragraph = section2->GetParagraphs()->GetItem(j);
			section1->GetBody()->GetChildObjects()->Insert(index + 1, dynamic_cast<Paragraph*>(paragraph->Clone()));
		}
	}

	//Save the document.
	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document1->Close();
	delete bn;
}

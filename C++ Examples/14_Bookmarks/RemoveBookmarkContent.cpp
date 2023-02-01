#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveBookmarkContent.docx";

	//Load the document from disk.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.            
	Bookmark* bookmark = document->GetBookmarks()->GetItem(L"Test");

	Paragraph* para = dynamic_cast<Paragraph*>(bookmark->GetBookmarkStart()->GetOwner());
	int startIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
	para = dynamic_cast<Paragraph*>(bookmark->GetBookmarkEnd()->GetOwner());
	int endIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

	//Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
	//This method is only to remove the content of the bookmark.
	for (int i = startIndex + 1; i < endIndex; i++)
	{
		para->GetChildObjects()->RemoveAt(startIndex + 1);
	}

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

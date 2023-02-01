#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveBookmark.docx";

	//Load the document from disk.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.
	Bookmark* bookmark = document->GetBookmarks()->GetItem(L"Test");

	//Remove the bookmark, not its content.
	document->GetBookmarks()->Remove(bookmark);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

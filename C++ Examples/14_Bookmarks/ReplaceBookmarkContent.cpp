#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceBookmarkContent.docx";

	//Load the document from disk.
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Locate the bookmark.
	BookmarksNavigator* bookmarkNavigator = new BookmarksNavigator(doc);
	bookmarkNavigator->MoveToBookmark(L"Test");

	//Replace the context with new.
	bookmarkNavigator->ReplaceBookmarkContent(L"This is replaced content.", false);

	//Save the document.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
	delete bookmarkNavigator;
}

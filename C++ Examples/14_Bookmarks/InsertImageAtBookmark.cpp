#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring imagePath = input_path + L"Word.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertImageAtBookmark.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create an instance of BookmarksNavigator
	BookmarksNavigator* bn = new BookmarksNavigator(doc);

	//Find a bookmark named Test
	bn->MoveToBookmark(L"Test", true, true);

	//Add a section
	Section* section0 = doc->AddSection();

	//Add a paragraph for the section
	Paragraph* paragraph = section0->AddParagraph();

	Image* image = Image::FromFile(imagePath.c_str());

	//Add a picture into the paragraph
	DocPicture* picture = paragraph->AppendPicture(image);

	//Add the paragraph at the position of bookmark
	bn->InsertParagraph(paragraph);

	//Remove the section0
	doc->GetSections()->Remove(section0);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
	delete bn;
}

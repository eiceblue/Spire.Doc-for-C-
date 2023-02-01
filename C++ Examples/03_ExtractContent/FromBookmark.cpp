#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FromBookmark.docx";

	//Create the source document
	Document* sourcedocument = new Document();

	//Load the source document from disk.
	sourcedocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	Document* destinationDoc = new Document();

	//Add a section for destination document
	Section* section = destinationDoc->AddSection();

	//Add a paragraph for destination document
	Paragraph* paragraph = section->AddParagraph();

	//Locate the bookmark in source document
	BookmarksNavigator* navigator = new BookmarksNavigator(sourcedocument);

	//Find bookmark by name
	navigator->MoveToBookmark(L"Test", true, true);

	//get text Body part
	TextBodyPart* textBodyPart = navigator->GetBookmarkContent();

	//Create a TextRange type list
	vector<TextRange*> list;

	//Traverse the items of text Body
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		auto item = textBodyPart->GetBodyItems()->GetItem(i);
		//if it is paragraph
		if (dynamic_cast<Paragraph*>(item) != nullptr)
		{
			//Traverse the ChildObjects of the paragraph
			for (int i = 0; i < (dynamic_cast<Paragraph*>(item))->GetChildObjects()->GetCount(); i++)
			{
				auto childObject = (dynamic_cast<Paragraph*>(item))->GetChildObjects()->GetItem(i);
				//if it is TextRange
				if (dynamic_cast<TextRange*>(childObject) != nullptr)
				{
					//Add it into list
					TextRange* range = dynamic_cast<TextRange*>(childObject);
					list.push_back(range);
				}
			}
		}
	}

	//Add the extract content to destinationDoc document
	for (int m = 0; m < list.size(); m++)
	{
		paragraph->GetChildObjects()->Add(list[m]->Clone());
	}

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourcedocument->Close();
	destinationDoc->Close();
	delete sourcedocument;
	delete destinationDoc;
	delete navigator;
}
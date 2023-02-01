#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Bookmark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyBookmarkContent.docx";
	
	//Load the document from disk.
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.
	Bookmark* bookmark = doc->GetBookmarks()->GetItem(L"Test");
	DocumentObject* docObj = nullptr;

	//Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
	//Then need to find its outermost parent object(Table),
	//and get the start/end index of current object on GetBody().
	if ((dynamic_cast<Paragraph*>(bookmark->GetBookmarkStart()->GetOwner()))->GetIsInCell())
	{
		docObj = bookmark->GetBookmarkStart()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
	}
	else
	{
		docObj = bookmark->GetBookmarkStart()->GetOwner();
	}
	int startIndex = doc->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->IndexOf(docObj);

	if ((dynamic_cast<Paragraph*>(bookmark->GetBookmarkEnd()->GetOwner()))->GetIsInCell())
	{
		docObj = bookmark->GetBookmarkEnd()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
	}
	else
	{
		docObj = bookmark->GetBookmarkEnd()->GetOwner();
	}
	int endIndex = doc->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->IndexOf(docObj);

	//Get the start/end index of the bookmark object on the paragraph.
	Paragraph* para = dynamic_cast<Paragraph*>(bookmark->GetBookmarkStart()->GetOwner());
	int pStartIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
	para = dynamic_cast<Paragraph*>(bookmark->GetBookmarkEnd()->GetOwner());
	int pEndIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

	//Get the content of current bookmark and copy.
	TextBodySelection* select = new TextBodySelection(doc->GetSections()->GetItem(0)->GetBody(), startIndex, endIndex, pStartIndex, pEndIndex);
	TextBodyPart* body = new TextBodyPart(select);
	for (int i = 0; i < body->GetBodyItems()->GetCount(); i++)
	{
		doc->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->Add((body->GetBodyItems())->GetItem(i)->Clone());

	}

	//Save the document.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
	delete body;
}

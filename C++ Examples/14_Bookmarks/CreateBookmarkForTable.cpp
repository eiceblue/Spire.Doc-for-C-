#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateBookmarkForTable.docx";

	//Create word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Create bookmark for a table
	CreateBookmarkForTable(document, section);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void CreateBookmarkForTable(Document* doc, Section* section)
{
	//Add a paragraph
	Paragraph* paragraph = section->AddParagraph();

	//Append text for added paragraph
	TextRange* txtRange = paragraph->AppendText(L"The following example demonstrates how to create bookmark for a table in a Word document.");

	//Set the font in italic
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Append bookmark start
	paragraph->AppendBookmarkStart(L"CreateBookmark");

	//Append bookmark end
	paragraph->AppendBookmarkEnd(L"CreateBookmark");

	//Add table
	Table* table = section->AddTable(true);

	//Set the number of rows and columns
	table->ResetCells(2, 2);

	//Append text for table cells		
	TextRange* range = table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"sampleA");
	range = table->GetRows()->GetItem(0)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"sampleB");
	range = table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"120");
	range = table->GetRows()->GetItem(1)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"260");

	//Get the bookmark by index.
	Bookmark* bookmark = doc->GetBookmarks()->GetItem(0);

	//Locate the bookmark by name.
	BookmarksNavigator* navigator = new BookmarksNavigator(doc);
	navigator->MoveToBookmark(bookmark->GetName());

	//Add table to TextBodyPart
	TextBodyPart* part = navigator->GetBookmarkContent();
	part->GetBodyItems()->Add(table);

	//Replace bookmark cotent with table
	navigator->ReplaceBookmarkContent(part);

	delete navigator;
}

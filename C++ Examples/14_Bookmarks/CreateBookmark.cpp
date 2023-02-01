#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateBookmark.docx";

	//Create word document.
	Document* document = new Document();

	//Create a new section.
	Section* section = document->AddSection();

	CreateBookmark(section);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void CreateBookmark(Section* section)
{
	Paragraph* paragraph = section->AddParagraph();
	TextRange* txtRange = paragraph->AppendText(L"The following example demonstrates how to create bookmark in a Word document.");
	txtRange->GetCharacterFormat()->SetItalic(true);

	section->AddParagraph();
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Simple Create Bookmark.");
	txtRange->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetCornflowerBlue());
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	//Write simple CreateBookmarks.
	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"SimpleCreateBookmark");
	paragraph->AppendText(L"This is a simple bookmark.");
	paragraph->AppendBookmarkEnd(L"SimpleCreateBookmark");

	section->AddParagraph();
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Nested Create Bookmark.");
	txtRange->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetCornflowerBlue());
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	//Write nested CreateBookmarks.
	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"Root");
	txtRange = paragraph->AppendText(L" This is Root data ");
	txtRange->GetCharacterFormat()->SetItalic(true);

	paragraph->AppendBookmarkStart(L"NestedLevel1");
	txtRange = paragraph->AppendText(L" This is Nested Level1 ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	txtRange->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetDarkSlateGray());

	paragraph->AppendBookmarkStart(L"NestedLevel2");
	txtRange = paragraph->AppendText(L" This is Nested Level2 ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	txtRange->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetDimGray());

	paragraph->AppendBookmarkEnd(L"NestedLevel2");
	paragraph->AppendBookmarkEnd(L"NestedLevel1");
	paragraph->AppendBookmarkEnd(L"Root");
}

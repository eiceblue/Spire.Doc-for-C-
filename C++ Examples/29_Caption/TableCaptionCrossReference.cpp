#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"TableCaptionCrossReference.docx";

	//Create word document
	Document* document = new Document();

	//Get the first section
	Section* section = document->AddSection();

	//Create a table
	Table* table = section->AddTable(true);
	table->ResetCells(2, 3);
	//Add caption to the table
	IParagraph* captionParagraph = table->AddCaption(L"Table", CaptionNumberingFormat::Number, CaptionPosition::BelowItem);

	//Create a bookmark
	wstring bookmarkName = L"Table_1";
	Paragraph* paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(bookmarkName.c_str());
	paragraph->AppendBookmarkEnd(bookmarkName.c_str());

	//Replace bookmark content
	BookmarksNavigator* navigator = new BookmarksNavigator(document);
	navigator->MoveToBookmark(bookmarkName.c_str());
	TextBodyPart* part = navigator->GetBookmarkContent();
	part->GetBodyItems()->Clear();
	part->GetBodyItems()->Add(captionParagraph);
	navigator->ReplaceBookmarkContent(part);

	//Create cross-reference field to point to bookmark "Table_1"
	Field* field = new Field(document);
	field->SetType(FieldType::FieldRef);
	field->SetCode(L"REF Table_1 \\p \\h");

	//Insert line breaks
	for (int i = 0; i < 3; i++)
	{
		paragraph->AppendBreak(BreakType::LineBreak);
	}

	//Insert field to paragraph
	paragraph = section->AddParagraph();
	TextRange* range = paragraph->AppendText(L"This is a table caption cross-reference, ");
	range->GetCharacterFormat()->SetFontSize(14);
	paragraph->GetChildObjects()->Add(field);

	//Insert FieldSeparator object
	FieldMark* fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
	paragraph->GetChildObjects()->Add(fieldSeparator);

	//Set display text of the field
	TextRange* tr = new TextRange(document);
	tr->SetText(L"Table 1");
	tr->GetCharacterFormat()->SetFontSize(14);
	tr->GetCharacterFormat()->SetTextColor(Color::GetDeepSkyBlue());
	paragraph->GetChildObjects()->Add(tr);

	//Insert FieldEnd object to mark the end of the field
	FieldMark* fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
	paragraph->GetChildObjects()->Add(fieldEnd);

	//Update fields
	document->SetIsUpdateFields(true);
	//Save the Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
	delete navigator;
}

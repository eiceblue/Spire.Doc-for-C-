#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateCrossReference.docx";

	//Create Word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Create a bookmark.
	Paragraph* paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"MyBookmark");
	paragraph->AppendText(L"Text inside a bookmark");
	paragraph->AppendBookmarkEnd(L"MyBookmark");

	//Insert line breaks.
	for (int i = 0; i < 4; i++)
	{
		paragraph->AppendBreak(BreakType::LineBreak);
	}

	//Create a cross-reference field, and link it to bookmark.                    
	Field* field = new Field(document);
	field->SetType(FieldType::FieldRef);
	field->SetCode(L"REF MyBookmark \\p \\h");

	//Insert field to paragraph.
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"For more information, see ");
	paragraph->GetChildObjects()->Add(field);

	//Insert FieldSeparator object.
	FieldMark* fieldSeparator = new FieldMark(document, FieldMarkType::FieldSeparator);
	paragraph->GetChildObjects()->Add(fieldSeparator);

	//Set display text of the field.
	TextRange* tr = new TextRange(document);
	tr->SetText(L"above");
	paragraph->GetChildObjects()->Add(tr);

	//Insert FieldEnd object to mark the end of the field.
	FieldMark* fieldEnd = new FieldMark(document, FieldMarkType::FieldEnd);
	paragraph->GetChildObjects()->Add(fieldEnd);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

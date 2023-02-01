#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateIFField.docx";

	//Create Word document.
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Add a new paragraph.
	Paragraph* paragraph = section->AddParagraph();

	// Define a method of creating an IF Field.
	CreateIfField(document, paragraph);

	//Define merged data.
	vector<LPCWSTR_S> fieldName = { L"Count" };
	vector<LPCWSTR_S> fieldValue = { L"2" };

	//Merge data into the IF Field.
	document->GetMailMerge()->Execute(fieldName, fieldValue);

	//Update all fields in the document.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;

}

void CreateIfField(Document* document, Paragraph* paragraph)
{
	IfField* ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");

	paragraph->GetItems()->Add(ifField);
	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");

	ParagraphBase* end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	FieldMark* fm = dynamic_cast<FieldMark*>(end);
	fm->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(dynamic_cast<FieldMark*>(end));
}
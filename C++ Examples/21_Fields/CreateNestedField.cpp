#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateNestedField.docx";

	//Create Word document.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = document->GetSections()->GetItem(0);

	Paragraph* paragraph = section->AddParagraph();

	//Create an IF field
	IfField* ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	//Create the embedded IF field
	IfField* ifField2 = new IfField(document);
	ifField2->SetType(FieldType::FieldIf);
	ifField2->SetCode(L"IF ");
	paragraph->GetChildObjects()->Add(ifField2);
	paragraph->GetItems()->Add(ifField2);
	paragraph->AppendText(L"\"200\" < \"50\"   \"200\" \"50\" ");
	IParagraphBase* embeddedEnd = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(dynamic_cast<FieldMark*>(embeddedEnd))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(embeddedEnd);
	ifField2->SetEnd(dynamic_cast<FieldMark*>(embeddedEnd));

	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");
	IParagraphBase* end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(dynamic_cast<FieldMark*>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(dynamic_cast<FieldMark*>(end));

	//Update all fields in the document.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
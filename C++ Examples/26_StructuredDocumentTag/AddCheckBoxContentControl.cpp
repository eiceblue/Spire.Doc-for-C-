#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCheckBoxContentControl.docx";

	//Create a document
	Document* document = new Document();

	//Add a new section.
	Section* section = document->AddSection();

	//Add a paragraph
	Paragraph* paragraph = section->AddParagraph();

	//Append textRange for the paragraph
	TextRange* txtRange = paragraph->AppendText(L"The following example shows how to add CheckBox content control in a Word document. \n");

	//Append textRange 
	txtRange = paragraph->AppendText(L"Add CheckBox Content Control:  ");

	//Set the font format
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Create StructureDocumentTagInline for document
	StructureDocumentTagInline* sdt = new StructureDocumentTagInline(document);

	//Add sdt in paragraph
	paragraph->GetChildObjects()->Add(sdt);

	//Specify the type
	sdt->GetSDTProperties()->SetSDTType(SdtType::CheckBox);

	//Set properties for control
	SdtCheckBox* scb = new SdtCheckBox();
	sdt->GetSDTProperties()->SetControlProperties(scb);

	//Add textRange format
	TextRange* tr = new TextRange(document);
	tr->GetCharacterFormat()->SetFontName(L"MS Gothic");
	tr->GetCharacterFormat()->SetFontSize(12);

	//Add textRange to StructureDocumentTagInline
	sdt->GetChildObjects()->Add(tr);

	//Set checkBox as checked
	scb->SetChecked(true);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

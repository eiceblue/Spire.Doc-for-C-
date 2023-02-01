#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertMergeField.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	Section* section = document->GetSections()->GetItem(0);

	Paragraph* par = section->AddParagraph();

	//Add merge field in the paragraph
	MergeField* field = dynamic_cast<MergeField*>(par->AppendField(L"MyFieldName", FieldType::FieldMergeField));

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

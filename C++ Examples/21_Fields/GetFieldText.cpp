#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetFieldText.txt";

	wstring* sb = new wstring();

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get all fields in document
	FieldCollection* fields = document->GetFields();
	for (int i = 0; i < fields->GetCount(); i++)
	{
		Field* field = fields->GetItem(0);
		//Get field text
		wstring fieldText = field->GetFieldText();
		sb->append(L"The field text is \"" + fieldText + L"\".\r\n");
	}
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << sb->c_str();
	out.close();
	document->Close();
	delete document;
	delete sb;
}

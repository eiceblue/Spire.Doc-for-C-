#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"MailMerge.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetMergeFieldName.txt";

	wstring* str = new wstring();

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get merge field name
	vector<LPCWSTR_S> fieldNames = document->GetMailMerge()->GetMergeFieldNames();

	str->append(L"The document has " + to_wstring(fieldNames.size()) + L" merge fields.");
	str->append(L" The below is the name of the merge field:\n");
	for (auto name : fieldNames)
	{
		str->append(name);
		str->append(L"\n");
	}

	wofstream out;
	out.open(outputFile.c_str());
	out.flush();
	out << str->c_str();
	out.close();
	document->Close();
	delete document;
	delete str;
}

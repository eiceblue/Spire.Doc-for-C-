#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CountWordsNumber.txt";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Count the number of words.
	wstring* content = new wstring();
	content->append(L"CharCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCount()));
	content->append(L"\n");
	content->append(L"CharCountWithSpace: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCountWithSpace()));
	content->append(L"\n");
	content->append(L"WordCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetWordCount()));

	//Save to file.
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << content->c_str();
	out.close();
	document->Close();
	delete document;
	delete content;
}

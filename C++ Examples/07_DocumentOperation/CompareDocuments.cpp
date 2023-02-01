#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring file1 = input_path + L"SupportDocumentCompare1.docx";
	wstring file2 = input_path + L"SupportDocumentCompare2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CompareDocuments.docx";

	//Load the first document
	Document* doc1 = new Document();
	doc1->LoadFromFile(file1.c_str());
	//Load the second document
	Document* doc2 = new Document();
	doc2->LoadFromFile(file2.c_str());
	//Compare the two documents
	doc1->Compare(doc2, L"E-iceblue");

	//Save as docx file.
	doc1->SaveToFile(outputFile.c_str(), Spire::Doc::FileFormat::Docx2013);
	doc1->Close();
	doc2->Close();
	delete doc1;
	delete doc2;
}

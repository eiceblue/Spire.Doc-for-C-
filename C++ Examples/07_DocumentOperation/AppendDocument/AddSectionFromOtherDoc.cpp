#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"SampleB_1.docx";
	wstring inputFile_2 = input_path + L"Sample_two sections.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddSectionFromOtherDoc.docx";

	//Open a Word document as target document
	Document* TarDoc = new Document(inputFile_1.c_str());
	//Open a Word document as source document
	Document* SouDoc = new Document(inputFile_2.c_str());
	//Get the second section from source document
	Section* Ssection = SouDoc->GetSections()->GetItem(1);

	//Add the section in target document
	TarDoc->GetSections()->Add(Ssection->Clone());

	//Save to file
	TarDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	SouDoc->Close();
	TarDoc->Close();
	delete TarDoc;
	delete SouDoc;
}

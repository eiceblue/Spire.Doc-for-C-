#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LockSpecifiedSections.docx";

	//Create Word document.
	Document* document = new Document();

	//Add new sections.
	Section* s1 = document->AddSection();
	Section* s2 = document->AddSection();

	//Append some text to section 1 and section 2.
	s1->AddParagraph()->AppendText(L"Spire.Doc demo, section 1");
	s2->AddParagraph()->AppendText(L"Spire.Doc demo, section 2");

	//Protect the document with AllowOnlyFormFields protection type.
	document->Protect(ProtectionType::AllowOnlyFormFields, L"123");

	//Unprotect section 2
	s2->SetProtectForm(false);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}

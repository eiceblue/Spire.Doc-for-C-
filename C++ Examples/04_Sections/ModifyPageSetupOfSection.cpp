#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_N2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ModifyPageSetupOfSection.docx";

	//Load Word from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Loop through all sections
	for (int i = 0; i < doc->GetSections()->GetCount(); i++)
	{
		Section* section = doc->GetSections()->GetItem(i);
		//Modify the margins
		section->GetPageSetup()->SetMargins(new MarginsF(100, 80, 100, 80));
		//Modify the page size
		section->GetPageSetup()->SetPageSize(PageSize::Letter());
	}

	// Or only modify one section
	// For example, modify the page setup of the first section
	//Section section0 = doc.Sections[0];
	//section0.PageSetup.Margins = new MarginsF(100, 80, 100, 80);
	//section0.PageSetup.FooterDistance = 35.4f;
	//section0.PageSetup.HeaderDistance = 34.4f;

	//Save the Word file
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}

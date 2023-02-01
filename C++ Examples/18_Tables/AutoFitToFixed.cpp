#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AutoFitToFixed.docx";

	//Create a document
	Document* document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	Section* section = document->GetSections()->GetItem(0);
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));
	//The table is set to a fixed size
	table->AutoFit(AutoFitBehaviorType::FixedColumnWidths);
	//Save to file
	document->SaveToFile(outputFile.c_str());
	document->Close();
	delete document;
}

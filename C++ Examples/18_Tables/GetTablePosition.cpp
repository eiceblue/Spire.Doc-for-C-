#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TableSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetTablePosition.txt";

	//Create a document
	Document* document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	//Get the first section
	Section* section = document->GetSections()->GetItem(0);
	//Get the first table
	Table* table = dynamic_cast<Table*>(section->GetTables()->GetItemInTableCollection(0));

	wstring* stringBuidler = new wstring();

	//Verify whether the table uses "Around" text wrapping or not.
	if (table->GetTableFormat()->GetWrapTextAround())
	{
		TablePositioning* positon = table->GetTableFormat()->GetPositioning();

		stringBuidler->append(L"Horizontal:");
		stringBuidler->append(L"Position:" + to_wstring(positon->GetHorizPosition()) + L" pt");
		//stringBuidler->append(L"Absolute Position:" + positon->GetHorizPositionAbs() + L", Relative to:" + positon->GetHorizRelationTo());
		stringBuidler->append(L"");
		stringBuidler->append(L"Vertical:");
		stringBuidler->append(L"Position:" + to_wstring(positon->GetVertPosition()) + L" pt");
		//stringBuidler->append(L"Absolute Position:" + positon->GetVertPositionAbs() + L", Relative to:" + positon->GetVertRelationTo());
		stringBuidler->append(L"");
		stringBuidler->append(L"Distance from surrounding text:");
		stringBuidler->append(L"Top:" + to_wstring(positon->GetDistanceFromTop()) + L" pt, Left:" + to_wstring(positon->GetDistanceFromLeft()) + L" pt");
		stringBuidler->append(L"Bottom:" + to_wstring(positon->GetDistanceFromBottom()) + L"pt, Right:" + to_wstring(positon->GetDistanceFromRight()) + L" pt");
	}

	//Save file.
	wofstream write(outputFile);
	write << stringBuidler->c_str();
	write.close();
	//File::WriteAllText(outputFile.c_str(), stringBuidler->toString());
	document->Close();
	delete document;
	delete stringBuidler;
}

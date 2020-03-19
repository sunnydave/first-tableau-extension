import React, {useState, useEffect} from "react";
import {Modal} from "react-bootstrap";
import LoadingIndicatorComponent from "./LoadingIndicator";
import SheetListComponent from "./SheetListComponent";
import DataTableComponent from "./DataTableComponent";
const {tableau} = window;

function Extension(){
    const [isLoading, setIsLoading] = useState(true);
    const [selectedSheet, setSelectedSheet] = useState(undefined);
    const [sheetNames, setSheetNames] = useState([]);
    const [rows, setRows] = useState([]);
    const [headers, setHeaders] = useState([]);

    let unregisterEventFn = undefined;

    useEffect(() => {
        tableau.extensions.initializeAsync().then(() => {
            const sheetNames = tableau.extensions.dashboardContent.dashboard.worksheets.map(worksheet => worksheet.name);
            setSheetNames(sheetNames);
            const selectedSheet = tableau.extensions.settings.get('sheet');
            setSelectedSheet(selectedSheet);

            const sheetSelected = !!selectedSheet;
            setIsLoading(sheetSelected);

            if(!!sheetSelected){
                loadSelectedMarks(selectedSheet);
            }
        })
    },[]);


    const getSelectedSheet = (sheet) => {
        const sheetName = sheet || selectedSheet;
        return tableau.extensions.dashboardContent.dashboard.worksheets.find(worksheet => worksheet.name === sheetName);
    };
    const loadSelectedMarks = (sheet) => {
        if(unregisterEventFn){
            unregisterEventFn();
        }

        const worksheet = getSelectedSheet(sheet);
        worksheet.getSelectedMarksAsync().then(marks => {
            const worksheetData = marks.data[0];
            const rows = worksheetData.data.map(row => row.map(cell => cell.value));
            const headers = worksheetData.columns.map(column => column.fieldName);
            setRows(rows);
            setHeaders(headers);
            setIsLoading(false);
        });
        unregisterEventFn = worksheet.addEventListener(tableau.TableauEventType.MarkSelectionChanged, () => {
            setIsLoading(true);
            loadSelectedMarks(sheet);
        })
    };

    const onSelectSheet = (sheet) => {
        tableau.extensions.settings.set('sheet',sheet);
        setIsLoading(true);
        tableau.extensions.settings.saveAsync().then(() => {
            setSelectedSheet(sheet);
            loadSelectedMarks();
        });
    };

    const mainContent = (rows.length > 0)
        ? (<DataTableComponent rows={rows} headers={headers}/>)
        : (<h4>No Data Found</h4>);

    let output = <di>{mainContent}</di>;
    if (isLoading) {
        output = <LoadingIndicatorComponent msg='Loading' />;
    }
    if(!selectedSheet){
        output =
            <Modal show>
                <Modal.Header>
                    <Modal.Title>Choose a Sheet</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    <SheetListComponent sheetNames={sheetNames} onSelectSheet={onSelectSheet}/>
                </Modal.Body>
            </Modal>
    };
    return (
        <div>{output}</div>
    );

}

export default Extension;

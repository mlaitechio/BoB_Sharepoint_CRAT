import React from 'react';
import PDF from 'react-pdf-watermark';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Dialog, DialogBody, DialogContent, DialogSurface, Spinner } from '@fluentui/react-components';
import { Constants } from '../../Constants/Constants';
import { Text } from '@microsoft/sp-core-library'

export interface IMyPDFViewerProps {
    pdfFilePath: any;
    watermarkText: any;
    currentSelectedFileContent?: any;
    context?: WebPartContext;
    documentLoaded: () => void;
}

export interface IMyPDFViewerState {
    pages: any;
    page: any;
    blobFile: any;

}

export default class MyPdfViewer extends React.Component<IMyPDFViewerProps, IMyPDFViewerState> {
    constructor(props) {
        super(props)

        this.state = {
            pages: 1,
            page: 1,
            blobFile: ``
        }
    }

    public async componentDidMount() {
        const { currentSelectedFileContent, context } = this.props;

        await this.waterMark_ConvertToBase64PDF(currentSelectedFileContent, `${this.props.watermarkText}`).then((val) => {
            this.setState({ blobFile: val }, () => {
                this.props.documentLoaded();
            })
        }).catch((error) => {
            console.log(error);
            this.props.documentLoaded();
        })
    }

    handlePrevious = () => {
        this.setState({ page: this.state.page - 1 });
    }

    handleNext = () => {
        this.setState({ page: this.state.page + 1 });
    }

    renderPagination = (page, pages) => {
        let previousButton = <li className="previous" onClick={this.handlePrevious}><a href="#"><i className="fa fa-arrow-left"></i> Previous</a></li>;
        if (page === 1) {
            previousButton = <li className="previous disabled"><a href="#"><i className="fa fa-arrow-left"></i> Previous</a></li>;
        }
        let nextButton = <li className="next" onClick={this.handleNext}><a href="#">Next <i className="fa fa-arrow-right"></i></a></li>;
        if (page === pages) {
            nextButton = <li className="next disabled"><a href="#">Next <i className="fa fa-arrow-right"></i></a></li>;
        }
        return (
            <nav>
                <ul className="pager">
                    {previousButton}
                    {nextButton}
                </ul>
            </nav>
        );
    }

    render() {
        const { blobFile } = this.state
        let pagination = null;
        if (this.state.pages) {
            pagination = this.renderPagination(this.state.page, this.state.pages);
        }
        return (
            <div>
                {/* {blobFile != null && <iframe src={blobFile} width={"100%"} height={"700px"} />} */}
                {/* {blobFile == null && this.workingOnIt()} */}
                {blobFile != null && <object data={blobFile} width={"100%"} height={"800px"} style={{ marginTop: -35 }}></object>}

            </div>
        )
    }

    private workingOnIt = (): JSX.Element => {

        let submitDialogJSX = <>

            <Dialog modalType="alert" defaultOpen={true}>
                <DialogSurface style={{ maxWidth: 250 }}>
                    <DialogBody style={{ display: "block" }}>
                        <DialogContent>
                            {<Spinner labelPosition="below" label={"Working on It..."}></Spinner>}
                        </DialogContent>
                    </DialogBody>
                </DialogSurface>
            </Dialog>

        </>;
        return submitDialogJSX;
    }

    private waterMark_ConvertToBase64PDF = async (fileContent, watermarkText) => {

        const pdfDoc = await PDFDocument.load(fileContent);
        const totalPages = pdfDoc.getPageCount();

        for (let pageNum = 0; pageNum < totalPages; pageNum++) {
            const page = pdfDoc.getPage(pageNum);
            const { width, height } = page.getSize();
            const textFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
            const fontSize = 100;

            const font = pdfDoc.embedStandardFont(StandardFonts.Helvetica);
            const headerHeight = 50;


            const textWidth = font.widthOfTextAtSize(Constants.infoPDFText, fontSize);
            const textHeight = font.heightAtSize(fontSize);

            const startX = (width - textWidth) / 2;
            const startY =
                height + headerHeight - (headerHeight - textHeight) / 2 - textHeight;

            page.moveTo(startX, startY);

            // page.drawText(Text.format(Constants.infoPDFText, new Date().toLocaleDateString()), {
            //     x: 5,
            //     y: 5,
            //     size: 12,
            //     font: font,
            //     opacity: 1,
            //     color: rgb(0.86, 0.09, 0.26),
            // });

            page.drawText(watermarkText, {
                x: width / 6,
                y: (1.6 * height) / 6,
                size: fontSize,
                font: textFont,
                opacity: 0.4,
                color: rgb(0.8392156862745098, 0.807843137254902, 0.792156862745098),
                rotate: degrees(30)
            });
        }

        let pdfBytes = await pdfDoc.save();

        let base64File = this.bufferToBase64(pdfBytes).then((val) => {
            const base64WithoutPrefix = val.substring('data:application/octet-stream;base64,'.length);

            const bytes = atob(base64WithoutPrefix);
            let length = pdfBytes.length;
            let out = new Uint8Array(length);

            while (length--) {
                out[length] = bytes.charCodeAt(length);
            }
            let blobFile = new Blob([out], { type: "application/pdf" });

            return Promise.resolve(URL.createObjectURL(blobFile))
        }).catch((error) => {
            return Promise.reject(error)
        });

        return base64File

    }


    private bufferToBase64 = async (buffer): Promise<any> => {
        // use a FileReader to generate a base64 data URI:
        const base64url = await new Promise(r => {
            const reader = new FileReader()
            reader.onload = () => r(reader.result)
            reader.readAsDataURL(new Blob([buffer]))
        });

        // remove the `data:...;base64,` part from the start
        return Promise.resolve(base64url);
    }
}


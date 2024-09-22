import React from 'react';
import PDF from 'react-pdf-watermark';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Dialog, DialogBody, DialogContent, DialogSurface, FluentProvider, Spinner } from '@fluentui/react-components';
import { Constants } from '../../Constants/Constants';
import * as pdfjsLib from "pdfjs-dist/build/pdf";
//import PDFImagePreview from './PDFImagePreview';
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://unpkg.com/pdfjs-dist@2.5.207/build/pdf.worker.min.js";


export interface IMyPDFViewerProps {
    pdfFilePath: any;
    watermarkText: any;
    currentSelectedFileContent?: any;
    context?: WebPartContext;
    documentLoaded: () => void;
    footerText?: string;
    footerTextColor?: any;
    mode?: any;
}

export interface IMyPDFViewerState {
    pages: any;
    page: any;
    blobFile: any;
    imgSrc: any;
    imgList?: any[]

}

export default class MyPdfViewer extends React.Component<IMyPDFViewerProps, IMyPDFViewerState> {

    private base64Data: any;

    constructor(props) {
        super(props)

        this.state = {
            pages: 1,
            page: 1,
            blobFile: ``,
            imgSrc: ``,
            imgList: []
        }
    }

    public async componentDidMount() {
        const { currentSelectedFileContent, context, watermarkText, footerText, footerTextColor } = this.props;

        console.log(currentSelectedFileContent);

        await this.waterMark_ConvertToBase64PDF(currentSelectedFileContent, `${watermarkText}`, footerText, footerTextColor).then((val) => {
            this.setState({ blobFile: val }, async () => {

                this.props.documentLoaded();

            })
        }).catch((error) => {
            console.log(error);
            this.props.documentLoaded();
        });
    }

    // handlePrevious = () => {
    //     this.setState({ page: this.state.page - 1 });
    // }

    // handleNext = () => {
    //     this.setState({ page: this.state.page + 1 });
    // }

    // renderPagination = (page, pages) => {
    //     let previousButton = <li className="previous" onClick={this.handlePrevious}><a href="#"><i className="fa fa-arrow-left"></i> Previous</a></li>;
    //     if (page === 1) {
    //         previousButton = <li className="previous disabled"><a href="#"><i className="fa fa-arrow-left"></i> Previous</a></li>;
    //     }
    //     let nextButton = <li className="next" onClick={this.handleNext}><a href="#">Next <i className="fa fa-arrow-right"></i></a></li>;
    //     if (page === pages) {
    //         nextButton = <li className="next disabled"><a href="#">Next <i className="fa fa-arrow-right"></i></a></li>;
    //     }
    //     return (
    //         <nav>
    //             <ul className="pager">
    //                 {previousButton}
    //                 {nextButton}
    //             </ul>
    //         </nav>
    //     );
    // }

    render() {
        const { blobFile, imgList } = this.state;
        const { mode } = this.props
        // let pagination = null;
        // if (this.state.pages) {
        //     pagination = this.renderPagination(this.state.page, this.state.pages);
        // }

        return (
            <div style={{
                maxHeight: 600, overflow: "scroll",
                WebkitOverflowScrolling: "touch",
                touchAction:"auto",
                msTouchAction:"auto",
                scrollbarWidth: "thin",
                border: imgList.length > 0 ? "1px solid #ccc" : "0px",
                overflowX: "hidden",

            }} onContextMenu={(e) => { e.preventDefault(); return false }}>
                {/* {blobFile != null  && < iframe id="contentFile" src={blobFile} width={"100%"} height={"700px"} style={{ marginTop: -35 }} />} */}
                {/* {blobFile == null && this.workingOnIt()} */}
                {/* {blobFile != null && <object data={blobFile} width={"100%"} height={"800px"} style={{ marginTop: -35 }}></object>} */}
                {/* {blobFile != null && <div style={{ overflow: 'scroll', height: 600 }}>
                    <MobilePDFReader url={blobFile} />
                </div>} */}

                {
                    imgList.length == 0 && <div style={{ textAlign: "center" }}>
                        <FluentProvider>
                            <Spinner label={`Working on it..`} labelPosition="below"></Spinner>
                        </FluentProvider>
                    </div>
                }

                {imgList && imgList.length > 0 && imgList.map((val) => {
                    return <div style={{ paddingTop: 10, paddingBottom: 10, textAlign: "center", background: "#cccccc6b" }}>
                        <img src={val} width={mode ? "90%" : "60%"} />
                    </div>
                })}

                {/* {blobFile != "" && <>
                    <PDFViewer document={{ base64: this.base64Data }}>

                    </PDFViewer>
                </>} */}

                {/* {blobFile != null && <>
                    <canvas id="pdfFile"></canvas>
                </>} */}

            </div>
        )
    }


    private isMobile = () => {

        const regex = /Mobi|Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i;
        return regex.test(navigator.userAgent);
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

    private waterMark_ConvertToBase64PDF = async (fileContent, watermarkText, footerText, footerTextColor) => {

        const pdfDoc = await PDFDocument.load(fileContent);
        const totalPages = pdfDoc.getPageCount();

        for (let pageNum = 0; pageNum < totalPages; pageNum++) {
            const page = pdfDoc.getPage(pageNum);
            const { width, height } = page.getSize();
            const textFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
            const fontSize = 65;

            const font = pdfDoc.embedStandardFont(StandardFonts.Helvetica);
            const headerHeight = 50;


            const textWidth = font.widthOfTextAtSize(Constants.infoPDFText, fontSize);
            const textHeight = font.heightAtSize(fontSize);

            const startX = (width - textWidth) / 2;
            const startY =
                height + headerHeight - (headerHeight - textHeight) / 2 - textHeight;

            page.moveTo(startX, startY);

            page.drawText(watermarkText, {
                x: width / 7.3,
                y: (1.6 * height) / 2.6,
                size: fontSize,
                font: textFont,
                opacity: 0.4,
                color: rgb(0.8392156862745098, 0.807843137254902, 0.792156862745098),
                rotate: degrees(30)
            });



            page.drawText(watermarkText, {
                x: width / 7.3,
                y: (1.6 * height) / 7.3,
                size: fontSize,
                font: textFont,
                opacity: 0.4,
                color: rgb(0.8392156862745098, 0.807843137254902, 0.792156862745098),
                rotate: degrees(30)
            });

            page.drawText(footerText, {
                x: 5,
                y: 5,
                size: 11,
                font: font,
                opacity: 1,
                color: footerTextColor ?? rgb(0.02, 0.02, 0.02) //rgb(0.02, 0.02, 0.02)//rgb(0.86, 0.09, 0.26),
            });

        }

        let pdfBytes = await pdfDoc.save();

        this.renderPage(pdfBytes).then((imgList) => {
            this.setState({ imgList })
        }).catch((error) => {
            console.log(error)
        })

        let base64File = this.bufferToBase64(pdfBytes).then((val) => {
            //const base64WithoutPrefix = val.substring('data:application/octet-stream;base64,'.length);
            const base64WithoutPrefix = val.substring('data:application/octet-stream;base64,'.length);
            this.base64Data = base64WithoutPrefix;
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


    private renderPage = async (data) => {
        // setLoading(true);
        const imagesList = [];
        const canvas = document.createElement("canvas");
        canvas.setAttribute("className", "canv");
        const pdf = await pdfjsLib.getDocument({ data }).promise;
        //const pdfDoc = await PDFDocument.load(fileContent);
        for (let i = 1; i <= pdf.numPages; i++) {
            var page = await pdf.getPage(i);
            var viewport = page.getViewport({ scale: 1.5 });
            canvas.height = viewport.height;
            canvas.width = viewport.width;
            var render_context = {
                canvasContext: canvas.getContext("2d"),
                viewport: viewport,
            };
            await page.render(render_context).promise;
            let img = canvas.toDataURL("image/png");
            imagesList.push(img);

        }

        return Promise.resolve(imagesList);

        // this.setState({ pdfImgList: imagesList })
        // setNumOfPages((e) => e + pdf.numPages);
        // setImageUrls((e) => [...e, ...imagesList]);
    };
}


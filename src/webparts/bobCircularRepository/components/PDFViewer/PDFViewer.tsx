import React from 'react';
import PDF from 'react-pdf-watermark';

export interface IMyPDFViewerProps {
    pdfFilePath: any;
    watermarkText: any
}

export interface IMyPDFViewerState {
    pages: any;
    page: any;
}

export default class MyPdfViewer extends React.Component<IMyPDFViewerProps, IMyPDFViewerState> {
    constructor(props) {
        super(props)

        this.state = {
            pages: 0,
            page: 0
        }
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
        let pagination = null;
        if (this.state.pages) {
            pagination = this.renderPagination(this.state.page, this.state.pages);
        }
        return (
            <div>
                <PDF
                    file={`https://pdf-lib.js.org/assets/with_update_sections.pdf`}
                    page={this.state.page}
                    watermark={this.props.watermarkText}
                    watermarkOptions={{
                        transparency: 0.5,
                        fontSize: 55,
                        fontStyle: 'Bold',
                        fontFamily: 'Arial'
                    }}
                    onDocumentComplete={() => { /* Do anything on document loaded like remove loading, etc */ }}
                    onPageRenderComplete={(pages, page) => this.setState({ page, pages })}
                />
                {pagination}
            </div>
        )
    }
}


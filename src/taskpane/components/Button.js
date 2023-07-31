import React from "react";
import { Button } from "@fluentui/react-components";
export class ButtonExample extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            text: "",
            content: [],
        };
    }

    postData = async (data) => {
        try {
            const response = await fetch("https://enagramm.com/API/SpellChecker/CheckTextTest", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({ Text: data, ToolsGlobalLangID: 1, ToolsLangCode: "GEO" }),
            });

            if (!response.ok) {
                throw new Error("Network response was not ok");
            }

            const responseData = await response.json();
            const correctWords = responseData.lstWR.filter((item) => item.Correct === 1).map((item) => item.Word);
            const incorrectWords = responseData.lstWR.filter((item) => item.Correct === 0).map((item) => item.Word);
            this.setState({ content: [responseData] }, async () => {
                console.log("Response from server:", responseData);

                await Word.run(async (context) => {
                    const body = context.document.body;
                    const paragraphs = body.paragraphs;
                    paragraphs.load("text, font");

                    await context.sync();

                    // Clear incorrect highlights
                    for (const paragraph of paragraphs.items) {
                        const text = paragraph.text;
                        const words = text.split(/\s+/);
                        for (const word of words) {
                            if (incorrectWords.includes(word)) {
                                paragraph.font.highlightColor = null;
                            }
                        }
                    }

                    // Highlight correct words
                    for (const paragraph of paragraphs.items) {
                        const text = paragraph.text;
                        const words = text.split(/\s+/);
                        for (const word of words) {
                            if (correctWords.includes(word)) {
                                const searchResults = paragraph.search(word, { matchCase: true });
                                context.load(searchResults, "font");
                                await context.sync();

                                for (const searchResult of searchResults.items) {
                                    searchResult.font.highlightColor = "yellow";
                                }
                            }
                        }
                    }
                });
            });
        } catch (error) {
            console.error("Error fetching data from server:", error);
        }
    };
    // if (correctWords.includes(word)) {
    //     const range = paragraph.getRange(wordIndex, word.length);
    //     range.track();
    //     paragraph.font.highlightColor = null; // Remove any existing highlight
    //     range.font.highlightColor = "yellow"; // Highlight the correct word
    // }
    handleWordClick = async (word) => {
        console.log(word);
        try {
            await Word.run(async (context) => {
                const body = context.document.body;
                const range = body.getRange();
                const searchResults = range.search(word, { matchCase: true });

                context.load(searchResults, "items");
                await context.sync();
                searchResults.items.forEach((searchResult) => {
                    searchResult.insertContentControl();

                });

                await context.sync();
            });
        } catch (error) {
            console.error("Error inserting content control:", error);
        }
    };


    getContentText = async () => {
        try {
            await Word.run(async (context) => {
                const body = context.document.body;
                context.load(body, "text");

                await context.sync();

                const contentText = body.text;
                if (contentText.length > 0) {
                    this.setState({ text: contentText });
                    console.log("Content text of the document:", contentText);

                    this.postData(contentText);
                }
            });
        } catch (error) {
            console.log("Error: " + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    };

    componentDidUpdate(prevProps, prevState) {
        if (this.state.text !== prevState.text) {
            this.postData(this.state.text.split(" "));
        }
    }

    render() {
        let { disabled } = this.props;
        const { content } = this.state;

        return (
            <div className="ms-BasicButtonExample">
                <br />
                <ul>
                    {content?.map((data, index) => {
                        return data.lstWR.map((item, index2) => {
                            if (item.Correct === 1) {
                                return (
                                    <li className="correct-word" key={index2}>
                                        <button onClick={() => this.handleWordClick(item.Word)}>{item.Word}</button>
                                    </li>
                                );
                            }
                        });
                    })}
                </ul>
                <Button appearance="primary" disabled={disabled} size="large" className="submit-btn" onClick={this.getContentText}>
                    Check
                </Button>
            </div>
        );
    }
}
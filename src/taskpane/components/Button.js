import * as React from "react"
import { Button, Label } from "@fluentui/react-components"

/* global Word */

export class ButtonExample extends React.Component {
    constructor(props, context) {
        super(props, context);
        this.state = {
            text: context,
            content: [],
            clickedWord: "", // To store the word clicked in the UI

        };
    }
    highlightText = async () => {
        const { clickedWord } = this.state;

        if (!clickedWord) {
            console.log("No word selected.");
            return;
        }

        // Office.js API to find and highlight the clicked word in the Word document
        await Word.run(async (context) => {
            const searchResults = context.document.body.search(clickedWord, { ignoreCase: false });

            // Load the search results
            context.load(searchResults, "font");
            // Synchronize the changes with the Word document
            console.log(context.document.body)
            await context.sync();
            // Apply highlight 
            searchResults.items.forEach((result) => {
                //     const isCorrect = content.some((data) =>
                //           data.lstWR.some((item) => item.Correct === 1 && item.Word.toLowerCase() === result.text.toLowerCase())
                // );
                // if (isCorrect) {
                //     const searchResults = body.search(words[i], { ignoreCase: true });
                //     searchResults.load("font");
                //     await context.sync();

                //     searchResults.items.forEach((result) => {
                //         result.font.highlightColor = "yellow";
                //     });
                // }
                result.font.highlightColor = "yellow";
            });

            console.log("Text highlighted:", clickedWord);
        });
    };
    handleResponseClick = (clickedWord) => {
        // Check if the clicked word is not an empty string
        if (clickedWord.trim() !== "") {
            // Update the state with the clicked word
            this.setState({ clickedWord }, () => {
                this.highlightText();
            });
        }
    };

    componentDidMount() {
        this.getContentText();
    }

    componentDidUpdate(prevProps, prevState) {
        if (this.state.text !== prevState.text) {
            this.postData(this.state.text);
        }
    }

    insertText = async () => {
        // Write text to the document when the button is selected.
        await Word.run(async context => {
            let body = context.document.body
            body.insertParagraph("Hello Fluent UI React!", Word.InsertLocation.end)
            await context.sync()
        })
    }

    postData = async (data) => {
        try {
            const response = await fetch("https://enagramm.com/API/SpellChecker/CheckTextTest", {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ "Text": data, "ToolsGlobalLangID": 1, "ToolsLangCode": "GEO" }),
            });

            if (!response.ok) {
                throw new Error('Network response was not ok');
            }

            return response.json();
        } catch (error) {
            console.error('Error:', error);
        }
    }

    // get context text
    getContentText = async () => {
        try {
            await Word.run(async context => {
                var body = context.document.body;
                context.load(body);

                // Execute the queued commands and return a promise
                await context.sync();

                // Get the text of the body and display it
                var contentText = body.text;
                if (contentText.length > 0) {

                    this.setState({ text: contentText });
                    console.log("Content text of the document:", contentText);

                    const data = await this.postData(contentText);
                    console.log("Response from server:", data);

                    this.setState({
                        content: [data],
                    });
                }
            });
        } catch (error) {
            console.log("Error: " + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    }
    render() {
        let { disabled } = this.props
        const { content } = this.state;

        return (
            <div className="ms-BasicButtonExample">
                <br />
                <ul>
                    {content?.map((data, index) => {
                        return (
                            data.lstWR.map((item, index2) => {
                                if (item.Correct === 1) {

                                    return (

                                        <li className="correct-word" key={index2}><button onClick={() => this.handleResponseClick(item.Word)}
                                        >{item.Word}</button></li>
                                    )
                                }
                            })

                        )
                    }
                    )}
                </ul>
                <Button
                    appearance="primary"
                    disabled={disabled}
                    size="large"
                    className="submit-btn"
                    onClick={this.getContentText}
                >
                    Check
                </Button>
            </div>
        )
    }
}

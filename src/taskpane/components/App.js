import * as React from "react";
import PropTypes from "prop-types";
import { ButtonExample } from './Button';

/* global Word, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  render() {
    return (
      <div className="ms-welcome">
        <ButtonExample />
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

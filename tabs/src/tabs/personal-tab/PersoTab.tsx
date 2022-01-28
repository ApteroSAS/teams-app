import React from "react";
import { Welcome } from "./Welcome";

var showFunction = Boolean(process.env.REACT_APP_FUNC_NAME);

export default function PersoTab() {
  return (
    <div>
        <Welcome showFunction={showFunction} />
    </div>
  );
}

import * as React from "react";
import { useState } from "react";

interface SearchProps{
    onSearch: (value:string) => void;
}

const Search : React.FC<SearchProps> = ({ onSearch }) => {

    const [search, setSearch] = useState("");

    const onInputChange = (value:string) => {
        setSearch(value);
        onSearch(value);
    };

    return(
            <div className="my-2">
            <input type="text" 
                className="form-control" 
                placeholder="Search"
                value={search}
                onChange={ e => onInputChange(e.target.value)}
            />
            </div>
    )

};

export default Search;
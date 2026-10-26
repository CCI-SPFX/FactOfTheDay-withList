import * as React from "react";
import { useEffect, useState } from "react";
import type { IFactOftheDayProps } from "./IFactOftheDayProps";
import "./style.css";

const FactOftheDay: React.FC<IFactOftheDayProps> = ({
  list,
  imgURL,
  siteUrl,
}) => {
  const [items, setItems] = useState<any[]>([]);

        const formatDate = (date: string) =>
  new Date(date).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "long",
    year: "numeric",
  }).replace(/ (\d{4})$/, ", $1");
  useEffect(() => {
    const fetchData = async () => {
      try {
        const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${list}')/items?$select=*`;

        const response = await fetch(apiUrl, {
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        });

        const data = await response.json();

        const mappedData = data.d.results.map((item: any) => ({
          title: item.Title,
          desc: item.Description,
          name: item.Name,
          date: item.Date,
        }));

        setItems(mappedData);
        console.log(mappedData);
        
      } catch (error) {
        console.error("Error fetching data:", error);
      }
    };

    fetchData();
  }, [siteUrl, list]);

  const firstItem = items.length > 0 ? items[0] : null;

  return (
    <div className="factCard">
      <div className="cardHeader">
        <h2>{firstItem?.title}</h2>
        <img src={imgURL} className="quoteIcon" alt="quote" />
      </div>

      {firstItem && (
        <>
          <p>{firstItem.desc}</p>

          <div className="metaData">
            <p>{firstItem.name}</p>
            <p>{formatDate(firstItem.date)}</p>
          </div>
        </>
      )}
    </div> 
  );
};

export default FactOftheDay;
import Button from "@/components/Button";

const DatabaseSavedWidget = ({ data, close }) => {
  return (
    <div className='z-100 border-3 fixed left-1/2 top-1/2 flex w-4/12 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center rounded-md border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider text-black'>
      Results of saving to database:
      <ol className="list-decimal">
        {data.map((item, idx) => (
          <li key={idx}>
            {!item.saved ? (
              <>
                {`Error saving ${
                  item.saveAttemptItem.title
                    ? item.saveAttemptItem.title
                    : "[Missing Title]"
                }:`}
                <ol type="a">
                  {item.errors.map((error, eIdx) => (
                    <li key={eIdx}>{error}</li>
                  ))}
                </ol>
              </>
            ) : (
              item.message
            )}
          </li>
        ))}
      </ol>
      <Button
        additionalStyling={"absolute right-2 top-2"}
        onClick={close}
        label={"Close"}
        width={100}
        fontSize={25}
      />
    </div>
  );
};

export default DatabaseSavedWidget;

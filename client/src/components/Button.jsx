const Button = ({ onClick, label, width, fontSize, additionalStyling }) => {
  console.log(width);
  return (
    <button
      style={{ width: `${width}px`, fontSize: `${fontSize}px` }}
      className={`${additionalStyling} cursor-pointer rounded-[8px] border-[3px] border-[var(--darkpink)] bg-[var(--lightpink)] pl-[4px] pr-[4px] font-["Just_Another_Hand"] tracking-wider text-black hover:bg-[var(--darkpink)]`}
      onClick={onClick}
    >
      {label}
    </button>
  );
};

export default Button;

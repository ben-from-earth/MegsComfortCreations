import CircularProgress from "@mui/material/CircularProgress";

const LoadingWidget = ({ message }) => {
  return (
    <div className='z-100 border-3 fixed left-1/2 top-1/2 flex h-1/4 w-1/4 -translate-x-1/2 -translate-y-1/2 flex-col content-center items-center justify-center rounded-md border-[var(--darkpink)] bg-[var(--lightpink)] p-2 font-["Just_Another_Hand"] text-2xl tracking-wider text-black'>
      <p>{message}</p>
      <CircularProgress sx={{ color: "#e1b3b5" }} />
    </div>
  );
};

export default LoadingWidget;

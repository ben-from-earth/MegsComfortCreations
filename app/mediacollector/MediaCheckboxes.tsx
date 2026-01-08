// // react, redux imports
// import { useDispatch } from 'react-redux';

// // necessary imports from collector state slice
// import {
//   mediaTypeDefinitions,
//   setChecks,
// } from 'lib/state/slices/collectorSlice';

// interface MediaCheckboxesProps {
//   mediaTypes: mediaTypeDefinitions[];
// }

// export default function MediaCheckboxes({ mediaTypes }: MediaCheckboxesProps) {
//   //setup connection to redux slice
//   const dispatch = useDispatch();

//   return (
//     <div className="m-6 flex flex-row content-center gap-5">
//       {mediaTypes.map(({ type, label, show }, idx) => (
//         <label key={type} className="text-3xl tracking-wider">
//           <input
//             checked={show}
//             className="m-1.5"
//             id={`${idx}`}
//             type="checkbox"
//             onChange={() => {
//               dispatch(setChecks(idx));
//             }}
//           />
//           {`${label}s`}
//         </label>
//       ))}
//     </div>
//   );
// }

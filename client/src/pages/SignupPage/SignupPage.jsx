import axios from 'axios';
import { Formik, Form, useField } from 'formik';
import { useState } from 'react';
import { useNavigate } from 'react-router';
import * as Yup from 'yup';

//server location import from .env
const serverDomain = import.meta.env.VITE_SERVER_DOMAIN;

const MyTextArea = ({ label, ...props }) => {
  const [field, meta] = useField(props);

  const showError = meta.touched && !!meta.error;
  return (
    <>
      <div className="flex w-full items-center gap-4 px-5">
        <label
          className='font-["Just_Another_Hand"] text-4xl'
          htmlFor={props.id || props.name}
        >
          {label}
        </label>
        <textarea
          className="ml-auto flex h-20 items-center justify-center rounded bg-white pl-2 pt-2"
          {...props}
          {...field}
          style={{
            borderColor: showError ? '1px solid red' : '1px solid white',
          }}
        />
      </div>
      {showError && <p className="text-lg text-red-600">{meta.error}</p>}
    </>
  );
};

const ItemForm = ({ handleSubmit, signupError }) => {
  return (
    <Formik
      initialValues={{ email: '', password: '', first_name: '', last_name: '' }}
      validationSchema={Yup.object({
        email: Yup.string()
          .trim()
          .email('Enter a valid email')
          .required('Email is required')
          .max(254, 'Email must be at most 254 characters'),
        password: Yup.string()
          .required('Password is required')
          .min(8, 'At least 8 characters')
          .max(128, 'At most 128 characters')
          .matches(/[A-Za-z]/, 'Include a letter')
          .matches(/\d/, 'Include a number'),
        first_name: Yup.string().required('First Name is required'),
        last_name: Yup.string().required('Last Name is required'),
      })}
      onSubmit={(values, { setSubmitting }) => {
        handleSubmit(values);
        setSubmitting(false);
      }}
    >
      <Form
        className={`border-3 tracking wider absolute left-1/2 top-1/2 flex h-fit w-fit -translate-x-1/2 -translate-y-1/2 flex-col items-center gap-2 rounded-md border-[var(--darkpink)] p-2 shadow-xl`}
      >
        <MyTextArea label="First Name:" name="first_name" type="text" />
        <MyTextArea label="Last Name:" name="last_name" type="text" />
        <MyTextArea label="Email:" name="email" type="text" />
        <MyTextArea label="Password:" name="password" type="text" />
        {signupError && (
          <p className="text-lg text-red-600">
            An account with this email already exists, try again.
          </p>
        )}
        <div className="buttonHolder">
          <button
            type="submit"
            className={`border-3 w-fit cursor-pointer rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-1 px-2 font-["Just_Another_Hand"] text-4xl tracking-wider text-black hover:bg-[var(--darkpink)]`}
          >
            Sign Up
          </button>
        </div>
      </Form>
    </Formik>
  );
};

const SignupPage = () => {
  const navigate = useNavigate();
  const [signupError, setSignupError] = useState(false);
  const handleSubmit = async (values) => {
    const res = await axios.post(
      `${serverDomain}/auth/signup`,
      {
        email: values.email,
        password: values.password,
        first_name: values.first_name,
        last_name: values.last_name,
      },
      { validateStatus: (status) => status < 500 },
    );
    const response = res.data;
    if (response.error) {
      setSignupError(true);
    } else if (response.user) {
      navigate('/Profile', { replace: true });
    }
  };
  return <ItemForm handleSubmit={handleSubmit} signupError={signupError} />;
};

export default SignupPage;

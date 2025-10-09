import { Formik, Form, useField } from 'formik';
import * as Yup from 'yup';
import axios from 'axios';
import { useState } from 'react';
import { useNavigate } from 'react-router';

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

const ItemForm = ({ handleSubmit, loginError }) => {
  return (
    <Formik
      initialValues={{ email: '', password: '' }}
      validationSchema={Yup.object({
        email: Yup.string()
          .trim()
          .email('Enter a valid email')
          .required('Email is required')
          .max(254, 'Email must be at most 254 characters'),
        password: Yup.string().required('Password is required'),
      })}
      onSubmit={(values, { setSubmitting }) => {
        handleSubmit(values);
        setSubmitting(false);
      }}
    >
      <Form className="flex flex-col items-center gap-3">
        <MyTextArea label="Email:" name="email" type="text" />
        <MyTextArea label="Password:" name="password" type="text" />
        {loginError && (
          <p className="text-lg text-red-600">Email or Password Invalid</p>
        )}
        <div className="buttonHolder">
          <button
            type="submit"
            className={`border-3 w-fit cursor-pointer rounded-lg border-[var(--darkpink)] bg-[var(--lightpink)] p-1 px-2 font-["Just_Another_Hand"] text-4xl tracking-wider text-black hover:bg-[var(--darkpink)]`}
          >
            Log In
          </button>
        </div>
      </Form>
    </Formik>
  );
};

const LoginPage = () => {
  const [loginError, setLoginError] = useState(false);
  const navigate = useNavigate();
  const handleSubmit = async (values) => {
    const res = await axios.post(
      `${serverDomain}/auth/login`,
      {
        email: values.email,
        password: values.password,
      },
      { validateStatus: (status) => status < 500, withCredentials: true },
    );
    const response = res.data;
    if (response.error) {
      setLoginError(true);
    } else if (res.data.user) {
      navigate('/Profile', { replace: true });
    }
  };
  return (
    <div className="border-3 tracking wider absolute left-1/2 top-1/2 flex h-fit w-fit -translate-x-1/2 -translate-y-1/2 flex-col items-center rounded-md border-[var(--darkpink)] p-2 shadow-xl">
      <h2 className='mb-4 font-["Just_Another_Hand"] text-4xl'>
        Please Log In!
      </h2>

      <ItemForm handleSubmit={handleSubmit} loginError={loginError} />
    </div>
  );
};

export default LoginPage;

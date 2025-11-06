// toastHelper.ts
import { toast, ToastOptions ,Id} from "react-toastify";
import 'react-toastify/dist/ReactToastify.css';

export type ToastType = "success" | "error" | "info" | "warning";

const defaultOptions: ToastOptions = {
  position: "top-right",
  autoClose: 3000,
  // autoClose: false,
  hideProgressBar: false,
  closeOnClick: true,
  pauseOnHover: true,
  draggable: true,
  theme: "colored"
};

const activeToasts: Record<string, Id> = {};
export const showToast = (type: ToastType, message: string, options: ToastOptions = {}) => {
  const toastOptions = { ...defaultOptions, ...options };

  // If a toast with the same message already exists
  if (activeToasts[message] && toast.isActive(activeToasts[message])) {
    // Blink or refresh effect: dismiss then re-show it
    toast.dismiss(activeToasts[message]);
    activeToasts[message] = toast[type](message, toastOptions);
    return;
  }

  // Otherwise, show a new toast
  const id = toast[type](message, toastOptions);
  activeToasts[message] = id;

  // switch (type) {
  //   case "success":
  //     toast.success(message, toastOptions);
  //     break;
  //   case "error":
  //     toast.error(message, toastOptions);
  //     break;
  //   case "info":
  //     toast.info(message, toastOptions);
  //     break;
  //   case "warning":
  //     toast.warning(message, toastOptions);
  //     break;
  //   default:
  //     toast(message, toastOptions);
  // }
};

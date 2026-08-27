"use client";

import { useEffect, useRef, useState } from "react";

const SPEECH_OUTPUT_START_EVENT = "jarvis-speech-output-start";
const SPEECH_OUTPUT_END_EVENT = "jarvis-speech-output-end";

function normalizeSpeechText(text: string) {
  return text
    .replace(/```(?:\w+)?\s*([\s\S]*?)```/g, "$1")
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/^\s{0,3}(?:#{1,6}|>|[-*+] |\d+\. )\s*/gm, "")
    .replace(/(\*\*|__|~~)(.*?)\1/g, "$2")
    .replace(/[*_]/g, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s*\n+\s*/g, ". ")
    .replace(/\s+/g, " ")
    .replace(/\.\s*\./g, ".")
    .trim();
}

export function useSpeechOutput() {
  const [isSupported, setIsSupported] = useState(false);
  const [isSpeaking, setIsSpeaking] = useState(false);
  const utteranceRef = useRef<SpeechSynthesisUtterance | null>(null);

  useEffect(() => {
    if (!("speechSynthesis" in window) || !("SpeechSynthesisUtterance" in window)) {
      return;
    }

    // Browser speech synthesis is only known after client mount.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setIsSupported(true);

    return () => {
      if (utteranceRef.current) {
        utteranceRef.current.onend = null;
        utteranceRef.current.onerror = null;
      }
      window.speechSynthesis.cancel();
      window.dispatchEvent(new Event(SPEECH_OUTPUT_END_EVENT));
      utteranceRef.current = null;
    };
  }, []);

  const stop = () => {
    if (!isSupported) return;

    if (utteranceRef.current) {
      utteranceRef.current.onend = null;
      utteranceRef.current.onerror = null;
    }
    window.speechSynthesis.cancel();
    window.dispatchEvent(new Event(SPEECH_OUTPUT_END_EVENT));
    utteranceRef.current = null;
    setIsSpeaking(false);
  };

  const speak = (text: string) => {
    if (!isSupported) return;

    const normalizedText = normalizeSpeechText(text);
    if (!normalizedText) return;

    if (utteranceRef.current) {
      utteranceRef.current.onend = null;
      utteranceRef.current.onerror = null;
    }
    window.speechSynthesis.cancel();

    const utterance = new SpeechSynthesisUtterance(normalizedText);
    utterance.onend = () => {
      if (utteranceRef.current !== utterance) return;
      utteranceRef.current = null;
      setIsSpeaking(false);
      window.dispatchEvent(new Event(SPEECH_OUTPUT_END_EVENT));
    };
    utterance.onerror = utterance.onend;
    utteranceRef.current = utterance;
    setIsSpeaking(true);
    window.dispatchEvent(new Event(SPEECH_OUTPUT_START_EVENT));
    window.speechSynthesis.speak(utterance);
  };

  return {
    isSupported,
    isSpeaking,
    speak,
    stop,
  };
}

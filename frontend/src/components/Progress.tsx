import React from 'react';

type Props = {
  progress: number;
};

const Progress: React.FC<Props> = ({ progress }) => (
  <div className="w-full bg-gray-200 rounded">
    <div
      className="bg-blue-600 text-white text-center text-xs font-semibold py-1 rounded"
      style={{ width: `${progress}%` }}
    >
      {progress}%
    </div>
  </div>
);

export default Progress;

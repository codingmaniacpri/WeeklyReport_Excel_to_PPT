import React, { type Dispatch, type SetStateAction } from 'react';

type Props = {
  companyName: string;
  setCompanyName: Dispatch<SetStateAction<string>>;
  projectTitle: string;
  setProjectTitle: Dispatch<SetStateAction<string>>;
  logoFile: File | null;
  setLogoFile: Dispatch<SetStateAction<File | null>>;
};

const Dashboard: React.FC<Props> = ({
  companyName, setCompanyName,
  projectTitle, setProjectTitle,
  logoFile, setLogoFile
}) => {

  const handleLogoChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setLogoFile(e.target.files[0]);
    }
  };

  return (
    <div className="mb-8 space-y-4 border p-4 rounded-md bg-gray-50">
      <div>
        <label className="block mb-1 font-semibold">Company Name</label>
        <input
          type="text"
          value={companyName}
          onChange={e => setCompanyName(e.target.value)}
          className="w-full border px-3 py-2 rounded"
          placeholder="Enter company name"
        />
      </div>

      <div>
        <label className="block mb-1 font-semibold">Project Title</label>
        <input
          type="text"
          value={projectTitle}
          onChange={e => setProjectTitle(e.target.value)}
          className="w-full border px-3 py-2 rounded"
          placeholder="Enter project title"
        />
      </div>

      <div>
        <label className="block mb-1 font-semibold">Upload Company Logo</label>
        <input
          type="file"
          accept="image/*"
          onChange={handleLogoChange}
          className="block"
        />
        {logoFile && <p className="mt-2 text-sm text-gray-700">{logoFile.name}</p>}
      </div>
    </div>
  );
};

export default Dashboard;

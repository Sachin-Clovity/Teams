import * as ResolverPkg from '@forge/resolver';
import { registerAdminResolvers } from './adminResolvers.js';
import { registerProjectResolvers } from './projectResolvers.js';
import { registerPersonalResolvers } from './personalResolvers.js';
import { registerIssuePanelResolvers } from './issuePanelResolvers.js';

const Resolver = ResolverPkg.default ?? ResolverPkg;
const resolver = new Resolver();

registerAdminResolvers(resolver);
registerProjectResolvers(resolver);
registerPersonalResolvers(resolver);
registerIssuePanelResolvers(resolver);

export default resolver;
